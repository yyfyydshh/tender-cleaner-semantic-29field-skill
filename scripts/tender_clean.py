# -*- coding: utf-8 -*-
"""招投标公告 29 字段语义清洗器。

支持 JSON、CSV、XLSX/XLSM 和 XLS 输入；优先以公告正文为事实来源，
不调用网络或 LLM，不补造公告未披露的信息。
"""
import argparse
import csv
import html
import json
import re
from datetime import date, datetime
from pathlib import Path

FIELDS = [
    '原标题', '来源链接', '项目名称', '招标编号', '项目执行地址', '项目规模',
    '招标范围', '项目工期', '招标金额', '获取招标文件时间', '投标截止时间',
    '开标时间', '开标方式', '开标地点', '招标人', '招标联系人', '招标人电话',
    '招标单位', '招标人邮箱', '招标人地址', '招标代理机构名称', '招标代理机构地址',
    '招标代理机构联系人', '招标代理机构电话', '招标代理机构邮箱',
    '中标人', '中标候选人', '中标金额', '中标公告时间'
]

COLUMN_ALIASES = {
    '标题': [
        '标题', '原标题', '公告标题', '项目标题', '项目名称',
        'title', 'name', 'project_name', 'projectname',
    ],
    '链接': [
        '链接', '来源链接', '公告链接', '详情链接', '原文链接',
        'url', 'link', 'href', 'source_url', 'sourceurl',
    ],
    '发布时间': [
        '发布时间', '发布日期', '公告时间', '公告日期', '发布日',
        'publish_time', 'publishtime', 'publish_date', 'publishdate', 'date',
    ],
    '文本': [
        '文本', '正文', '公告正文', '公告内容', '详情正文', '详情内容',
        'content', 'body', 'text', 'description', 'htmltext',
    ],
}

def clean_text(text):
    """清理文本"""
    text = '' if text is None else str(text)
    text = re.sub(r'(?i)<br\s*/?>|</p\s*>|</li\s*>|</div\s*>', '\n', text)
    text = re.sub(r'<[^>]+>', '', text)
    text = html.unescape(text)
    text = re.sub(r'\t+', ' ', text)
    text = re.sub(r' {2,}', ' ', text)
    return text.strip()


def structure_text(text):
    """给被网页抽取器压成单行的公告补回常见标签边界。"""
    labels = (
        r'项目编号|采购编号|招标编号|项目名称|采购方式|项目预算金额|项目预算|'
        r'采购预算|(?<!项目)预算金额|预算控制价|最高限价(?:（如有）)?|采购需求|合同履行期限|'
        r'供货地点|交货地点|供货时间|服务期限|实施周期|实施地点|招标内容|招标范围|'
        r'获取竞争性磋商文件|提交响应文件截止时间|截止时间、开启时间|开标时间|'
        r'采购人信息名称|采购代理机构信息名称|招标代理机构|采购代理机构|'
        r'项目联系人|(?<!项目)联系人|联系电话|联系方式|地\s*址|电\s*话|电子邮件'
    )
    # 只在“标签：”前补换行，不拆普通正文中的同名词。
    return re.sub(rf'(?<!\n)(?=(?:{labels})\s*[：:])', '\n', text)

def _flex_label(label):
    """将标签转为灵活正则（允许嵌入空格），如 '地 址' -> '地\\s*址'"""
    return re.sub(r'\s+', r'\\s*', label)

def extract_label(text, label, multiline=True):
    """
    提取标签后的值，支持多行。
    返回 (value, end_pos) 或 ('', -1)
    """
    flex = _flex_label(label)
    # 单行模式：标签：值 或 标签单独一行在值上方
    m = re.search(rf'{flex}[：:]\s*(.+?)(?:\n|\r|$)', text)
    if m:
        return m.group(1).strip(), m.end()
    
    if multiline:
        # 多行模式：标签单独一行，值在下一行
        m = re.search(rf'{flex}\s*\n\s*(.+?)(?:\n|\r|$)', text)
        if m:
            return m.group(1).strip(), m.end()
    
    return '', -1

def extract_all_labels(text, labels, multiline=True):
    """提取多个标签，返回匹配到的所有 (label, value) 对"""
    results = []
    for label in labels:
        val, _ = extract_label(text, label, multiline)
        if val:
            results.append((label, val))
    return results

def extract_contact_info(text):
    """
    提取完整联系人信息。
    策略：先划分"采购人段"和"代理机构段"，再分别提取。
    """
    result = {
        '招标人': '', '招标联系人': '', '招标人电话': '',
        '招标单位': '', '招标人邮箱': '', '招标人地址': '',
        '招标代理机构名称': '', '招标代理机构地址': '',
        '招标代理机构联系人': '', '招标代理机构电话': '', '招标代理机构邮箱': ''
    }
    
    # ===== 第一步：划分段落 =====
    # 采购人段标记
    tender_markers = [
        r'\d+\.\s*采购人信息', r'六、凡对本次[采招].+?提出询问',
        r'八、对本次招标提出询问', r'六、凡对本次公告', r'采购人',
    ]
    # 代理机构段标记
    agent_markers = [
        r'\d+\.\s*采购代理机构信息', r'2\.\s*采购代理机构信息',
        r'招标代理机构', r'采购代理机构', r'代理机构',
    ]
    
    # 找到文本中所有代理机构段和采购人段的分界
    tender_start = -1
    agent_start = -1
    
    for marker in tender_markers:
        m = re.search(marker, text)
        if m:
            tender_start = m.start()
            break
    
    for marker in agent_markers:
        m = re.search(marker, text)
        if m:
            agent_start = m.start()
            break
    
    # 如果没找到明确的采购人段，在末尾找"采购人"/"招标人"
    if tender_start < 0:
        m = re.search(r'(?:采购人|招标人)(?:名称)?[：:]', text)
        if m:
            tender_start = m.start()
    
    # 确定各段的文本范围
    # tender_text: 从采购人段开始到代理机构段（或文末）
    if tender_start >= 0:
        if agent_start > tender_start:
            tender_text = text[tender_start:agent_start]
        else:
            tender_text = text[tender_start:]
    else:
        tender_text = ''
    
    # agent_text: 从代理机构段开始到下一个大段落（或文末）
    if agent_start >= 0:
        agent_text = text[agent_start:]
    else:
        agent_text = ''
    
    # ===== 第二步：从采购人段提取 =====
    if tender_text:
        # 采购人名称 — 兼容 "名 称" "名称" 等变体
        m = re.search(r'(?:采购人|招标人)(?:名称)?[：:]\s*(.+?)(?:\n|$)', tender_text)
        if m:
            name = m.group(1).strip()
            if len(name) > 2:
                result['招标人'] = name
                result['招标单位'] = name
        
        if not result['招标人']:
            val, _ = extract_label(tender_text, r'名\s*称')
            if val and len(val) > 2 and '代理' not in val:
                result['招标人'] = val
                result['招标单位'] = val
        
        # 联系人 — 如果包含"电话"需要切分
        m = re.search(r'(?:联系人|项目联系人|报名及(?:论证)?联系人)[：:]\s*(.+?)(?:\n|$)', tender_text)
        if m:
            contact_val = m.group(1).strip()
            # 拆分 "孙老师 联系电话：0871-63629514"
            split_m = re.match(r'(.+?)[，,、\s]+(?:联系)?电[话話][：:]\s*(.+)', contact_val)
            if split_m:
                result['招标联系人'] = split_m.group(1).strip()
                if not result['招标人电话']:
                    result['招标人电话'] = split_m.group(2).strip()
            elif '电话' not in contact_val and '电 话' not in contact_val:
                result['招标联系人'] = contact_val
        
        # 电话/联系方式
        for label in [r'联系电话', r'联系方式', r'电\s*话']:
            m = re.search(rf'{label}[：:]\s*(.+?)(?:\n|$)', tender_text)
            if m:
                phone = m.group(1).strip()
                if not result['招标人电话']:
                    result['招标人电话'] = phone
                break
        
        # 地址 — 兼容 "地 址"
        m = re.search(r'地\s*址[：:]\s*(.+?)(?:\n|$)', tender_text)
        if m:
            addr = m.group(1).strip()
            if len(addr) > 3 and '送达' not in addr:
                result['招标人地址'] = addr
        
        # 邮箱（从采购人段）— ASCII模式只匹配英文
        emails = re.findall(r'[a-zA-Z0-9._\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}', tender_text)
        if emails:
            result['招标人邮箱'] = emails[0]
    
    # ===== 第三步：从代理机构段提取 =====
    if agent_text:
        # 代理机构名称 — 兼容 "名 称"
        m = re.search(r'名\s*称[：:]\s*(.+?)(?:\n|$)', agent_text)
        if m:
            name = m.group(1).strip()
            if len(name) > 3:
                result['招标代理机构名称'] = name
        
        # 代理机构地址 — 兼容 "地 址"
        m = re.search(r'地\s*址[：:]\s*(.+?)(?:\n|$)', agent_text)
        if m:
            addr = m.group(1).strip()
            if len(addr) > 5 and '送达' not in addr:
                result['招标代理机构地址'] = addr
        
        # 代理机构联系人
        m = re.search(r'项目联系人[：:]\s*(.+?)(?:\n|$)', agent_text)
        if m:
            contact_val = m.group(1).strip()
            split_m = re.match(r'(.+?)[，,、\s]+(?:联系)?电[话話][：:]\s*(.+)', contact_val)
            if split_m:
                result['招标代理机构联系人'] = split_m.group(1).strip()
                if not result['招标代理机构电话']:
                    result['招标代理机构电话'] = split_m.group(2).strip()
            elif '电话' not in contact_val:
                result['招标代理机构联系人'] = contact_val
        
        # 代理机构电话 — 兼容 "联系方式" "电 话"
        for label in [r'联系方式', r'电\s*话']:
            m = re.search(rf'{label}[：:]\s*(.+?)(?:\n|$)', agent_text)
            if m:
                phone = m.group(1).strip()
                if not result['招标代理机构电话']:
                    result['招标代理机构电话'] = phone
                break
        
        # 代理机构邮箱
        emails = re.findall(r'[a-zA-Z0-9._\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}', agent_text)
        if emails:
            result['招标代理机构邮箱'] = emails[0]
    
    # ===== 第四步：全文语义回退（不使用任何项目或机构硬编码） =====
    # 从整文本补全联系人和电话
    if not result['招标联系人']:
        m = re.search(r'联系人[：:]\s*(.+?)(?:\n|$)', text)
        if m:
            name = m.group(1).strip()
            pos = m.start()
            # 拆分联系人中的电话
            split_m = re.match(r'(.+?)[，,、\s]+(?:联系)?电[话話][：:]\s*(.+)', name)
            if split_m:
                name = split_m.group(1).strip()
                phone_from_name = split_m.group(2).strip()
            
            if agent_start >= 0 and pos > agent_start:
                if not result['招标代理机构联系人']:
                    result['招标代理机构联系人'] = name
                if split_m and not result['招标代理机构电话']:
                    result['招标代理机构电话'] = phone_from_name
            else:
                result['招标联系人'] = name
                if split_m and not result['招标人电话']:
                    result['招标人电话'] = phone_from_name
    
    if not result['招标人电话']:
        for label in [r'联系电话', r'联系方式', r'电\s*话']:
            m = re.search(rf'{label}[：:]\s*(\d[\d\-\(\)\s、，,\.]{6,})', text)
            if m:
                phone = m.group(1).strip()
                pos = m.start()
                if agent_start >= 0 and pos > agent_start:
                    if not result['招标代理机构电话']:
                        result['招标代理机构电话'] = phone
                else:
                    result['招标人电话'] = phone
                break
    
    # 清理电话中混入的姓名前缀，同时把姓名保存为联系人
    for phone_key, contact_key in [('招标人电话', '招标联系人'), ('招标代理机构电话', '招标代理机构联系人')]:
        if result[phone_key]:
            # 模式: "郭璐（0871）63614496" -> 提取"郭璐"为联系人，保留"（0871）63614496"为电话
            name_m = re.match(r'([\u4e00-\u9fff]{2,4})\s*[（(]', result[phone_key])
            if name_m:
                extracted_name = name_m.group(1)
                # 保持电话只保留括号部分
                result[phone_key] = result[phone_key][name_m.end()-1:]
                # 如果联系人还没设置，把提取到的名字填上
                if not result[contact_key]:
                    result[contact_key] = extracted_name
            # 如果值包含明显的姓名模式，只保留数字部分
            elif re.search(r'[\u4e00-\u9fff]', result[phone_key]):
                # 先尝试提取姓名
                name_part = re.match(r'([\u4e00-\u9fff]{2,4})', result[phone_key])
                digits = re.findall(r'[\d\-\(\)\s、，,\.]+', result[phone_key])
                if digits:
                    result[phone_key] = ''.join(digits).strip()
                if name_part and not result[contact_key]:
                    result[contact_key] = name_part.group(1)
    
    # 从整文本中提取邮箱（如果还没找到）
    all_emails = re.findall(r'[a-zA-Z0-9._\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}', text)
    if not result['招标人邮箱'] and all_emails:
        # 优先取靠近文末的邮箱（联系人信息通常在末尾）
        # 策略：取全文最后一个匹配的邮箱（通常紧邻采购人信息）
        result['招标人邮箱'] = all_emails[-1]

    # ===== 第五步：单行网页的语义回退与值清理 =====
    def clean_entity(value):
        value = re.split(
            r'(?:\d+(?:\.\d+)+\s*)?(?:地\s*址|联系人|联系电话|联系方式|'
            r'电\s*话|电子邮件|邮编|监督部门名称)[：:]?',
            value or '',
            maxsplit=1,
        )[0]
        value = re.split(
            r'(?<=[\u4e00-\u9fff])\d+(?:\.\d+)+\s*',
            value,
            maxsplit=1,
        )[0]
        return value.strip('。；，、 -')

    tender_patterns = [
        r'(?:项目业主|招标人)\s*为\s*([^，。\n]{2,100})',
        r'受\s*([^，。\n]{2,100}?)\s*(?:的)?委托',
        r'(?:采购人|招标人)(?:信息)?(?:名称)?[：:]\s*([^\n]{2,120})',
        r'采购人信息.{0,50}?名\s*称[：:]\s*([^\n]{2,120})',
    ]
    for pattern in tender_patterns:
        m = re.search(pattern, text, re.S)
        if m:
            name = clean_entity(m.group(1))
            if 2 < len(name) <= 100 and '代理' not in name:
                result['招标人'] = name
                result['招标单位'] = name
                break

    agent_patterns = [
        r'(?:招标|采购)代理机构[：:]\s*([^\n]{2,120})',
        r'([^\n，。]{2,100}?(?:有限责任公司|有限公司|事务所|中心))\s*受\s*[^\n，。]{2,100}?委托',
        r'采购代理机构信息.{0,50}?名\s*称[：:]\s*([^\n]{2,120})',
    ]
    for pattern in agent_patterns:
        m = re.search(pattern, text, re.S)
        if m:
            name = clean_entity(m.group(1))
            if 3 < len(name) <= 100:
                result['招标代理机构名称'] = name
                break

    phone_pattern = re.compile(
        r'(?:\+?86[-\s]?)?'
        r'(?:1[3-9]\d{9}|[（(]?0\d{2,3}[）)]?[-\s]?\d{7,8}(?:-\d{1,6})?|\d{7,8})'
    )

    if not result['招标人电话']:
        nearby = list(re.finditer(
            r'(?:报名[^\n：:]{0,20})?联系人[\s\S]{0,120}?'
            r'(?:联系电话|电\s*话)[：:]\s*([^\n。；]{7,100})',
            text,
        ))
        if nearby:
            result['招标人电话'] = nearby[-1].group(1)

    for phone_key in ['招标人电话', '招标代理机构电话']:
        phones = []
        for phone in phone_pattern.findall(result[phone_key] or ''):
            phone = re.sub(r'\s+', '', phone)
            if phone not in phones:
                phones.append(phone)
        result[phone_key] = '、'.join(phones[:5])

    # 宁可留空，也不把整段正文误当成实体字段。
    limits = {
        '招标人': 100, '招标单位': 100, '招标代理机构名称': 100,
        '招标联系人': 60, '招标代理机构联系人': 60,
        '招标人地址': 180, '招标代理机构地址': 180,
    }
    for key, limit in limits.items():
        if len(result[key]) > limit:
            result[key] = ''
    
    return result


def extract_date_range(text):
    """提取日期范围"""
    patterns = [
        r'请于\s*(\d{4}年\d{1,2}月\d{1,2}日)\s*至\s*(\d{4}年\d{1,2}月\d{1,2}日)',
        r'(\d{4}年\d{1,2}月\d{1,2}日)\s*至\s*(\d{4}年\d{1,2}月\d{1,2}日)',
        r'时间[：:]\s*(\d{4}年\d{1,2}月\d{1,2}日)\s*至\s*(\d{4}年\d{1,2}月\d{1,2}日)',
        r'请于\s*(\d{4}-\d{1,2}-\d{1,2})[^\n]{0,30}?至\s*(\d{4}-\d{1,2}-\d{1,2})',
        r'(\d{4}-\d{1,2}-\d{1,2})[^\n]{0,30}?至\s*(\d{4}-\d{1,2}-\d{1,2})',
    ]
    for p in patterns:
        m = re.search(p, text)
        if m:
            return f"{m.group(1)} 至 {m.group(2)}"
    return ''


def resolve_relative_deadline(relative_text, publish_date_str):
    """
    解析相对时间，如"发布公告后5个工作日内" → 从发布日推算出绝对日期。
    publish_date_str 格式如 "2026-07-15" 或 "2026年07月15日"
    返回格式化后的日期字符串，无法解析返回空。
    """
    from datetime import datetime, timedelta
    
    if not relative_text or not publish_date_str:
        return ''
    
    # 解析发布时间
    pub = None
    for fmt in ['%Y-%m-%d', '%Y年%m月%d日', '%Y.%m.%d']:
        try:
            pub = datetime.strptime(publish_date_str.strip(), fmt)
            break
        except ValueError:
            continue
    if not pub:
        return ''
    
    # 解析相对时间描述
    # 模式1: "发布公告后N个工作日内" / "发布公告后N个工作日内"
    m = re.search(r'(?:发布)?公告后\s*(\d+)\s*个?\s*工作日内?', relative_text)
    if m:
        days = int(m.group(1))
        return _add_working_days(pub, days)
    
    # 模式2: "自本公告发布之日起N个工作日" / "自本公告发布之日起N天"
    m = re.search(r'自(?:本)?公告发布之日起\s*(\d+)\s*个?\s*(工作日内?|天[内]?)', relative_text)
    if m:
        days = int(m.group(1))
        if '工作' in m.group(2):
            return _add_working_days(pub, days)
        else:
            return (pub + timedelta(days=days)).strftime('%Y年%m月%d日')
    
    # 模式3: "之日起N天内" / "之日起N天内"
    m = re.search(r'之日起\s*(\d+)\s*天内?', relative_text)
    if m:
        days = int(m.group(1))
        return (pub + timedelta(days=days)).strftime('%Y年%m月%d日')
    
    # 模式4: "公告后/发布后N天/日" 精简版
    m = re.search(r'(?:公告|发布|之日)\s*后?\s*(\d+)\s*[天日](?:内)?', relative_text)
    if m:
        days = int(m.group(1))
        return (pub + timedelta(days=days)).strftime('%Y年%m月%d日')
    
    return ''


def _add_working_days(start_date, days):
    """从 start_date 起加 N 个工作日"""
    from datetime import timedelta
    current = start_date
    remaining = days
    while remaining > 0:
        current += timedelta(days=1)
        if current.weekday() < 5:  # Mon=0, Fri=4
            remaining -= 1
    return current.strftime('%Y年%m月%d日')


def extract_datetime(text):
    """提取日期时间，兼容多种格式：如 2026年7月17日9:30 / 09点30分 / 10时00分"""
    if not text:
        return ''
    
    # 先统一 "时" → "点" 方便后续匹配
    text_normalized = re.sub(r'(\d{1,2})时(\d{2})分', r'\1点\2分', text)
    # 也处理 "10时00" 无"分"的情况 → "10:00"
    text_normalized = re.sub(r'(\d{1,2})时(\d{2})(?!分)', r'\1:\2', text_normalized)
    
    patterns = [
        r'(\d{4}年\d{1,2}月\d{1,2}日)\s*(\d{1,2}[:：]\d{2})',
        r'(\d{4}年\d{1,2}月\d{1,2}日)\s*(\d{1,2})点(\d{2})分',
        r'(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2})',
        r'(\d{4}年\d{1,2}月\d{1,2}日)',
    ]
    
    for p in patterns:
        m = re.search(p, text_normalized)
        if m:
            date = m.group(1)
            # 优先检查 3-group 模式（如 "点...分"），再检查 2-group
            if len(m.groups()) >= 3 and m.group(3):
                time_part = f"{m.group(2)}:{m.group(3)}"
            elif len(m.groups()) >= 2 and m.group(2):
                time_part = m.group(2).replace('：', ':')
            else:
                return date
            return f"{date} {time_part}"
    
    return ''


def chinese_num_to_int(cn):
    """中文数字字符串 → 整数。如 '贰拾陆万贰仟柒佰捌拾柒' → 262787"""
    UNIT_MAP = {'十': 10, '百': 100, '千': 1000, '万': 10000, '亿': 100000000,
                '拾': 10, '佰': 100, '仟': 1000}
    DIGIT_MAP = {'零': 0, '一': 1, '二': 2, '两': 2, '三': 3, '四': 4,
                 '五': 5, '六': 6, '七': 7, '八': 8, '九': 9,
                 '壹': 1, '贰': 2, '叁': 3, '肆': 4,
                 '伍': 5, '陆': 6, '柒': 7, '捌': 8, '玖': 9}
    section = 0    # 当前段（万以下累加）
    result = 0     # 亿/万级别累加
    cur_digit = 0  # 当前数字（等待单位）
    
    for ch in cn:
        if ch in DIGIT_MAP:
            cur_digit = DIGIT_MAP[ch]
        elif ch in UNIT_MAP:
            unit = UNIT_MAP[ch]
            if unit >= 10000:  # 万、亿
                if cur_digit > 0:
                    section += cur_digit
                    cur_digit = 0
                if section == 0:
                    section = 1
                result = (result + section) * unit
                section = 0
            else:  # 十、百、千
                if cur_digit == 0:
                    cur_digit = 1  # "十" 单独出现 = 10
                section += cur_digit * unit
                cur_digit = 0
        # 忽略其他字符（如 零 在 十百千 处理中会自然跳过添加）
    
    section += cur_digit
    return result + section


def extract_chinese_amount_full(text):
    """提取中文大写金额并转换为数字，如 贰拾陆万贰仟柒佰捌拾柒元叁角肆分 → 262787.34"""
    m = re.search(r'[零一二三四五六七八九十百千万亿两壹贰叁肆伍陆柒捌玖拾佰仟]+元[零一二三四五六七八九壹贰叁肆伍陆柒捌玖角分整]+', text)
    if not m:
        return 0.0
    s = m.group()
    yuan_part, *rest = s.split('元', 1)
    jiao_fen = rest[0] if rest else ''
    
    total = float(chinese_num_to_int(yuan_part))
    if '整' in jiao_fen:
        return total
    
    jiao, fen = 0, 0
    jm = re.search(r'([零一二三四五六七八九壹贰叁肆伍陆柒捌玖]+)角', jiao_fen)
    if jm:
        jiao = chinese_num_to_int(jm.group(1))
    fm = re.search(r'([零一二三四五六七八九壹贰叁肆伍陆柒捌玖]+)分', jiao_fen)
    if fm:
        fen = chinese_num_to_int(fm.group(1))
    return total + jiao * 0.1 + fen * 0.01


def _try_parse_amount(val_str, has_wan=False):
    """尝试从金额字符串解析数字，支持万元转换。返回 (number_or_None, is_ambiguous)"""
    val = val_str.strip('。；， ')
    # 检测万单位
    detected_wan = '万' in val
    has_wan = has_wan or detected_wan
    
    # 清理
    val = val.replace('¥', '').replace('￥', '').replace('元', '').replace('万', '').replace('，', '')
    val = val.replace(',', '').strip()
    
    # 如果还有括号注释，去掉
    val = re.sub(r'[\(（].*?[\)）]', '', val)
    # 只保留数字和点
    val = re.sub(r'[^\d.]', '', val)
    
    if not val:
        return None, False
    
    try:
        num = float(val)
        if has_wan:
            # 如果数字本身不大（< 100000）且有万标记 → 乘以10000
            # 如果数字很大（如1,592,000）且带万 → 可能原文格式错误，不乘
            if num < 10000:
                num *= 10000
            # 如果 num >= 10000 且有万标记 → 可能是如"1592000.00万"这种错误格式
            # 不额外乘，因为数字本身就很大了
        if num > 1000:
            return num, False
    except ValueError:
        pass
    return None, False


def extract_amount(text):
    """仅从明确的预算/限价/控制价语境提取金额，统一转换为元。"""
    labels = r'(?:预算金额|项目预算|采购预算|最高限价|预算控制价|最高控制价|招标控制价)'
    pattern = re.compile(
        rf'{labels}\s*(?:（\s*(?:人民币|含税)?\s*(?:元|万元)\s*）)?\s*(?:为|：|:)?\s*'
        r'(?:(?:人民币|含税|约)\s*)*[¥￥]?\s*([\d,]+(?:\.\d+)?)\s*(万)?\s*元',
        re.I,
    )
    for match in pattern.finditer(text):
        number = float(match.group(1).replace(',', ''))
        if match.group(2):
            number *= 10000
        context = text[max(0, match.start() - 20):match.end() + 20]
        if re.search(r'文件售价|标书费|保证金|代理费', context):
            continue
        return f'{number:.2f}'
    return ''


def clean_project_name(value):
    """清理标题/标签中的公告类型后缀，并保留项目本体名称。"""
    value = re.sub(r'^\s*[A-Za-z0-9][A-Za-z0-9._/-]*\s*[：:]\s*', '', value or '')
    value = re.sub(r'^\s*【[^】]+】\s*', '', value)
    suffixes = (
        r'(?:公开招标公告|竞争性磋商(?:招标)?公告|竞争性谈判公告|询比采购公告|'
        r'询比公告|询价公告|比选公告|遴选公告|单一来源采购公告|'
        r'咨询邀请公告|公开咨询公告|咨询公告|'
        r'更正公告|中标结果公告|成交结果公告|中标公告|成交公告|废标公告|'
        r'终止公告|招标公告|公告)\s*$'
    )
    return re.sub(suffixes, '', value).strip('。；，、 -')


def _clean_field_value(value, limit=300):
    """压缩字段值，避免正文或下一章节粘入结构化字段。"""
    value = re.sub(r'\s+', ' ', value or '').strip('。；，、 ')
    value = re.split(r'\s*(?:\d+、|\d+[.．]|[一二三四五六七八九十]+、)\s*', value, maxsplit=1)[0]
    return value[:limit].rstrip('。；，、 ')


def _extract_first_label(text, labels, limit=300):
    for label in labels:
        # 只接受“字段：值”，避免把目录标题后的下一段正文误认成字段值。
        value, _ = extract_label(text, label, multiline=False)
        if value:
            value = _clean_field_value(value, limit)
            if value:
                return value
    return ''


def extract_document_window(text):
    """只从“文件获取/售卖/发售”上下文提取时间，避免误取其他日期范围。"""
    patterns = [
        r'(?:招标|采购|竞争性磋商)文件(?:的)?(?:获取|发售|售卖)(?:时间|截止时间)?\s*(?:为|：|:)??\s*([^\n]{0,140})',
        r'请于\s*([^\n]{0,120}?\s*(?:至|到|—|-)\s*[^\n]{0,80})\s*(?:获取|购买|下载)(?:招标|采购|竞争性磋商)文件',
    ]
    for pattern in patterns:
        m = re.search(pattern, text, re.I)
        if not m:
            continue
        value = _clean_field_value(m.group(1), 160)
        date_span = re.search(
            r'(\d{4}[年-]\d{1,2}[月-]\d{1,2}(?:日)?(?:\s*\d{1,2}(?:[:：时点]\d{1,2})?(?:分)?)?\s*(?:至|到|—|-)\s*\d{4}[年-]\d{1,2}[月-]\d{1,2}(?:日)?(?:\s*\d{1,2}(?:[:：时点]\d{1,2})?(?:分)?)?)',
            value,
        )
        if date_span:
            return date_span.group(1)
        single_date = re.search(r'(\d{4}[年-]\d{1,2}[月-]\d{1,2}日?(?:\s*\d{1,2}(?:[:：时点]\d{1,2})?(?:分)?)?)', value)
        if single_date:
            return f'截至{single_date.group(1)}'
    return extract_date_range(text)


def extract_semantic_scale(text):
    """从工程/物资公告中提取可核验的规模描述。"""
    direct = re.findall(r'(?:总建筑面积|建筑面积|占地面积|需求数量)\s*[：:]\s*([^。；\n]{3,120})', text)
    if direct:
        return '；'.join(_clean_field_value(item, 120) for item in direct[:3])

    values = []
    for label in [r'项目规模', r'工程规模', r'建设规模']:
        value, _ = extract_label(text, label, multiline=False)
        value = _clean_field_value(value, 180)
        if value and value not in values:
            values.append(value)
    if values:
        return '；'.join(values[:3])

    # 标签缺失时，仅提取正文中明确的“总建筑面积/计划规模”短句。
    return ''


def enrich_contact_info(text, result):
    """从公告末尾联系方式区补全采购人、组织单位、联系人、电话和地址。"""
    tail = text[max(0, len(text) - 3000):]

    def last_value(pattern, limit=180):
        matches = list(re.finditer(pattern, tail, re.I))
        if not matches:
            return ''
        return _clean_field_value(matches[-1].group(1), limit)

    def credible_entity(value):
        if not value or re.search(r'^(?:执行标准|包件号|物资名称|账号|户名|开户行)', value):
            return False
        return not re.search(r'(?:银行|证书|平台技术支持)$', value)

    tenderer = last_value(r'(?:采购人|招标人)(?:名称)?\s*[：:]\s*([^\n]{2,180})')
    if credible_entity(tenderer) and not result['招标人']:
        result['招标人'] = tenderer
    if credible_entity(tenderer) and not result['招标单位']:
        result['招标单位'] = tenderer

    unit = last_value(r'(?:采购组织单位|招标组织单位|发布单位)\s*[：:]\s*([^\n]{2,180})')
    if unit:
        result['招标单位'] = unit

    address = last_value(r'地址\s*[：:]\s*([^\n]{4,220})', 220)
    if address and not result['招标人地址'] and '送达' not in address:
        result['招标人地址'] = address



def trim_section_tail(value):
    """截掉单行网页里黏在字段值后的下一个编号章节。"""
    value = re.split(r'(?<=。)\s*\d+(?:\.\d+)+\s+', value or '', maxsplit=1)[0].strip()
    return re.sub(r'\s*\d+(?:\.\d+)+\s*$', '', value).rstrip('。；，、 ')


def extract_project_number(text, title=''):
    """提取业务项目编号；存在平台 I 编号和业务编号时优先业务编号。"""
    candidates = re.findall(
        r'(?:项目|采购|招标)编号\s*[：:]\s*([A-Za-z0-9][A-Za-z0-9._/()（）-]{2,})',
        text,
    )
    candidates = [c.strip('。；，、()（） ') for c in candidates]
    # 处理“ZTZY-26-01”与下一章节“1.1.2”无空格黏连的网页文本。
    candidates = [re.sub(r'\d(?:\.\d+){2,}$', '', c) for c in candidates]
    business = [c for c in candidates if not re.fullmatch(r'I\d{15,}(?:G\d+)?', c)]
    if business:
        return business[0]
    if candidates:
        return candidates[0]

    m = re.search(r'[（(]([A-Za-z0-9][A-Za-z0-9._/-]{2,})[）)]', title)
    return m.group(1) if m else ''


def extract_procurement_text(record):
    """对单条记录按公告语义提取29字段。"""
    title = record.get('标题', '')
    link = record.get('链接', '')
    text = structure_text(clean_text(record.get('文本', '')))
    
    result = {f: '' for f in FIELDS}
    result['原标题'] = title
    result['来源链接'] = link
    if not result['来源链接']:
        # 只记录正文中明确出现的 URL；不把平台名称或首页猜成公告详情链接。
        m = re.search(r'https?://[^\s<>()（），。]+', text)
        if m:
            result['来源链接'] = m.group(0).rstrip('。；，、')
    
    is_scrap = '废标' in title
    is_consult = '咨询' in title
    is_win = '中标' in title or '成交' in title
    
    # === 项目名称 ===
    val = _extract_first_label(text, [r'项目名称', r'工程名称'])
    if val and len(val) <= 200:
        result['项目名称'] = clean_project_name(trim_section_tail(val))
    
    # 咨询类公告：项目名在 "开展XXX(采购)项目" 中
    if not result['项目名称'] and '咨询' in title:
        m = re.search(r'(?:开展|采购)(.+?(?:采购)?项目)', text)
        if m:
            result['项目名称'] = m.group(1).strip()
    
    # 咨询类：从标题提取（去掉"咨询邀请公告"等后缀）
    if not result['项目名称']:
        clean_title = clean_project_name(title)
        if len(clean_title) > 10 and clean_title != title:
            result['项目名称'] = clean_title
    
    # 最后兜底：从标题去掉公告类型后缀
    if not result['项目名称']:
        name_from_title = clean_project_name(title)
        if len(name_from_title) > 5:
            result['项目名称'] = name_from_title
    
    # === 招标编号 ===
    result['招标编号'] = extract_project_number(text, title) or record.get('编号', '')
    
    # === 项目执行地址 ===
    for label in [r'交货地点', r'★交货地点', r'项目实施地点', r'工程地点', r'建设地点', r'施工地点', r'服务地点', r'项目所在地']:
        val, _ = extract_label(text, label)
        if val:
            result['项目执行地址'] = _clean_field_value(val, 220)
            break
    if not result['项目执行地址'] and record.get('地区'):
        result['项目执行地址'] = record['地区']
    
    # === 项目规模 ===
    result['项目规模'] = extract_semantic_scale(text)
    
    # === 招标范围 ===
    result['招标范围'] = _extract_first_label(text, [r'招标范围', r'采购范围', r'招标内容', r'采购内容'], 300)
    if not result['招标范围']:
        m = re.search(r'(?:现对|对)([^。\n]{3,180}?)(?:进行)?(?:公开)?(?:招标|采购|选商)', text)
        if m:
            result['招标范围'] = _clean_field_value(m.group(1), 300)
    if re.search(r'检验报告|须提供|营业范围|质量保证|供货业绩|资格要求', result['招标范围']):
        result['招标范围'] = ''
    
    # === 项目工期 ===
    for label in [
        r'计划工期', r'供货期', r'供货时间', r'服务期限', r'实施周期', r'工期',
        r'★合同履行期限', r'合同履行期限', r'合同履约期限', r'★交货期',
    ]:
        val, _ = extract_label(text, label)
        if val:
            result['项目工期'] = trim_section_tail(val).strip('。；， ')
            break
    
    # === 招标金额 ===
    result['招标金额'] = extract_amount(text)
    
    # === 获取招标文件时间 ===
    result['获取招标文件时间'] = extract_document_window(text)
    
    # 咨询类：用"自公告之日起至"的时间范围
    if is_consult and not result['获取招标文件时间']:
        m = re.search(r'自公告之日(?:起|)(?:至|到)\s*(\d{4}年\d{1,2}月\d{1,2}日\s*\d{1,2}:\d{2})', text)
        if m:
            result['获取招标文件时间'] = extract_datetime(m.group(1))
    
    # === 投标截止时间 ===
    deadline_labels = [
        r'提交投标文件截止时间', r'提交响应文件截止时间',
        r'投标截止时间', r'投标截至时间', r'递交投标文件截止时间', r'投标文件递交截止时间',
        r'截止时间[、,，]?\s*开标时间', r'截止时间',
    ]
    for label in deadline_labels:
        val, _ = extract_label(text, label)
        if val:
            dt = extract_datetime(val)
            if dt:
                result['投标截止时间'] = dt
                break
    
    # 从项目概况中找 "于XXXX前提交"
    if not result['投标截止时间']:
        m = re.search(r'于\s*(\d{4}年\d{1,2}月\d{1,2}日\s*\d{1,2}[:：点时]\d{2}[分]?)', text)
        if m:
            result['投标截止时间'] = extract_datetime(m.group(1))
    
    # 从 "提交投标文件截止时间、开标时间和地点" 复合标题中找
    if not result['投标截止时间']:
        m = re.search(r'提交投标文件截止时间.*?(\d{4}年\d{1,2}月\d{1,2}日\s*\d{1,2}[:：点时]\d{2}[分]?)', text)
        if m:
            result['投标截止时间'] = extract_datetime(m.group(1))
    
    # 废标/嵌入式中找: "至X年X月X日X:XX止" / "至X年X月X日X:XX递交"
    if not result['投标截止时间']:
        m = re.search(r'(?:递交|响应文件|投标文件|磋商文件).*?(?:截止时间)?\s*[:：]?\s*(\d{4}年\d{1,2}月\d{1,2}日\s*\d{1,2}[:：点时]\d{2}[分]?)', text)
        if m:
            result['投标截止时间'] = extract_datetime(m.group(1))
    if not result['投标截止时间']:
        m = re.search(r'至\s*(\d{4}年\d{1,2}月\d{1,2}日\s*\d{1,2}[:：点时]\d{2}[分]?)\s*(?:止|递交|前)', text)
        if m:
            result['投标截止时间'] = extract_datetime(m.group(1))
    if not result['投标截止时间']:
        m = re.search(
            r'投标文件递交的截止时间(?:（[^）]*）)?\s*为\s*'
            r'(\d{4}-\d{1,2}-\d{1,2}\s+\d{1,2}:\d{2}(?::\d{2})?)',
            text,
        )
        if m:
            result['投标截止时间'] = extract_datetime(m.group(1))
    if not result['投标截止时间']:
        m = re.search(
            r'截止时间、(?:响应文件)?开启时间[：:]\s*'
            r'(\d{4}年\d{1,2}月\d{1,2}日\s*\d{1,2}[点时:：]\d{2}(?:分)?)',
            text,
        )
        if m:
            result['投标截止时间'] = extract_datetime(m.group(1))
    if not result['投标截止时间'] and record.get('截止时间'):
        result['投标截止时间'] = record['截止时间']
    
    # 咨询邀请公告：用报名/响应截止时间充当投标截止时间
    if is_consult and not result['投标截止时间']:
        # 模式1: 报名时间自公告之日起至XXXX
        m = re.search(r'报名时间自公告之日起至\s*(\d{4}年\d{1,2}月\d{1,2}日\s*\d{1,2}:\d{2})', text)
        if m:
            result['投标截止时间'] = extract_datetime(m.group(1))
        
        # 模式2: 报名截止时间/响应文件截止时间/递交截止时间/材料提交截止时间
        if not result['投标截止时间']:
            for label in [r'报名截止时间', r'响应文件截止时间', r'递交截止时间', r'材料提交截止时间']:
                val, _ = extract_label(text, label)
                if val:
                    dt = extract_datetime(val)
                    if dt:
                        result['投标截止时间'] = dt
                        break
        
        # 模式3: 相对时间 "发布公告后N个工作日内" — 用发布时间推算
        if not result['投标截止时间']:
            m = re.search(r'发布公告后\s*\d+\s*个?\s*工作日内?', text)
            if m:
                pub_date = record.get('发布时间', '')
                resolved = resolve_relative_deadline(m.group(), pub_date)
                if resolved:
                    result['投标截止时间'] = resolved
        
        # 模式4: "至XXX" 日期+时间
        if not result['投标截止时间']:
            m = re.search(r'至\s*(\d{4}年\d{1,2}月\d{1,2}日)\s*(\d{1,2}:\d{2})', text)
            if m:
                result['投标截止时间'] = f"{m.group(1)} {m.group(2)}"
    
    # === 开标时间 ===
    for label in [r'开标时间', r'开启时间']:
        val, _ = extract_label(text, label)
        if val:
            dt = extract_datetime(val)
            if dt:
                result['开标时间'] = dt
                break
    
    if not result['开标时间']:
        m = re.search(r'(?:电子)?(?:响应|投标)?文件(?:开启|解密)时间[：:]?\s*([^\n]{0,80})', text)
        if m:
            result['开标时间'] = extract_datetime(m.group(1))

    if not result['开标时间'] and result['投标截止时间']:
        # 竞争性磋商：截止时间=开标时间
        if '开启时间' in text and '截止时间及开启时间' in text:
            result['开标时间'] = result['投标截止时间']
    
    # 咨询邀请公告：开标时间 = 投标截止时间（论证会时间）
    if is_consult and not result['开标时间'] and result['投标截止时间']:
        result['开标时间'] = result['投标截止时间']
    
    # 咨询类：用投标截止时间反推获取招标文件时间（从发布时间到截止时间）
    if is_consult and not result['获取招标文件时间'] and result['投标截止时间']:
        pub_date = record.get('发布时间', '')
        if pub_date:
            pub_str = pub_date.strip()
            from datetime import datetime
            for fmt in ['%Y-%m-%d', '%Y年%m月%d日', '%Y.%m.%d']:
                try:
                    pub_dt = datetime.strptime(pub_str, fmt)
                    pub_str = pub_dt.strftime('%Y年%m月%d日')
                    break
                except ValueError:
                    continue
            result['获取招标文件时间'] = f"{pub_str} 至 {result['投标截止时间']}"
        else:
            m = re.search(r'发布公告后\s*\d+\s*个?\s*工作日内?', text)
            if m:
                result['获取招标文件时间'] = m.group()
    
    # === 开标方式 ===
    val, _ = extract_label(text, r'开标方式')
    if val:
        result['开标方式'] = val
    else:
        if '网上远程开标' in text:
            result['开标方式'] = '网上远程开标'
        elif '远程解密' in text:
            result['开标方式'] = '远程解密'
        elif re.search(r'电子(?:投标|响应)文件|线上递交|云链平台', text):
            result['开标方式'] = '线上电子开标/解密'
    
    # === 开标地点 ===
    for label in [r'开标地点', r'磋商地点', r'提交响应文件地点及磋商地点', r'投标地点[（(]网址[)）]']:
        val, _ = extract_label(text, label)
        if val and len(val) > 3:
            result['开标地点'] = val.strip('。；， ')
            break
    
    if not result['开标地点']:
        val, _ = extract_label(text, r'地点')
        if val and len(val) > 4 and '政采云' not in val:
            result['开标地点'] = val.strip('。；， ')

    if not result['开标地点']:
        platform = re.search(r'(?:中国铁建云链平台|电子交易平台)[（(]?([^）)\n]{0,100})', text)
        if platform:
            result['开标地点'] = _clean_field_value(platform.group(0), 160)

    if '中国铁建云链平台' in result['开标地点']:
        platform = re.search(r'中国铁建云链平台(?:[（(]([^）)\s]+)[）)])?', result['开标地点'])
        if platform:
            suffix = f'（{platform.group(1)}）' if platform.group(1) else ''
            result['开标地点'] = f'中国铁建云链平台{suffix}'
    elif len(result['开标地点']) > 160:
        result['开标地点'] = _clean_field_value(result['开标地点'], 160)
    
    # === 联系人信息 ===
    enrich_contact_info(text, result)
    contact = extract_contact_info(text)
    for k, v in contact.items():
        if not v or result[k]:
            continue
        if k in {'招标人', '招标单位'} and re.search(r'^(?:执行标准|包件号|物资名称|账号|户名|开户行|交通银行)', v):
            continue
        if k in {'招标联系人', '招标代理机构联系人'} and re.search(r'\d{7,}', v):
            continue
        if k in {'招标人电话', '招标代理机构电话'}:
            phones = re.findall(r'(?<!\d)(?:1[3-9]\d{9}|[（(]?0\d{2,3}[）)]?[-\s]?\d{7,8}(?:-\d{1,6})?)(?!\d)', v)
            v = '、'.join(dict.fromkeys(re.sub(r'\s+', '', phone) for phone in phones))
        if v:
            result[k] = v
    
    # === 中标信息 ===
    if is_scrap:
        result['中标人'] = '废标'
        val, _ = extract_label(text, r'废标原因')
        if val:
            result['中标候选人'] = val
    
    if is_win:
        val, _ = extract_label(text, r'中标人')
        if val:
            result['中标人'] = val
        val, _ = extract_label(text, r'中标金额')
        if val:
            result['中标金额'] = val
        val, _ = extract_label(text, r'中标候选人')
        if val:
            result['中标候选人'] = val
    
    # === 最终清理 ===
    for k in result:
        if isinstance(result[k], str):
            result[k] = re.sub(r'\s+', ' ', result[k]).strip('。；， ')
    
    return result


def _normalize_header(value):
    """归一化列名，兼容中英文、空格、下划线和常见标点。"""
    value = '' if value is None else str(value)
    return re.sub(r'[\s_\-—–:：/\\|（）()\[\]【】]+', '', value).lower()


def _as_text(value):
    """将表格单元格稳定转换为字符串。"""
    if value is None:
        return ''
    if isinstance(value, datetime):
        if value.hour == value.minute == value.second == value.microsecond == 0:
            return value.strftime('%Y-%m-%d')
        return value.strftime('%Y-%m-%d %H:%M:%S')
    if isinstance(value, date):
        return value.strftime('%Y-%m-%d')
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _first_alias(row, aliases):
    normalized = {_normalize_header(key): _as_text(value) for key, value in row.items()}
    for alias in aliases:
        value = normalized.get(_normalize_header(alias), '')
        if value:
            return value
    return ''


def _best_text_value(row):
    """从多列正文/文本中选择信息量最大的一列，而非按列名顺序取第一列。"""
    aliases = {_normalize_header(alias) for alias in COLUMN_ALIASES['文本']}
    candidates = []
    for key, value in row.items():
        normalized = _normalize_header(key)
        if not value:
            continue
        if normalized in aliases or any(normalized.startswith(alias) for alias in aliases if len(alias) >= 2):
            # 正文和详情正文优先；其他同类列（如“文本1/文本2”）作为补充。
            priority = 2 if normalized in {_normalize_header('正文'), _normalize_header('公告正文'), _normalize_header('详情正文')} else 1
            candidates.append((priority, len(value), key, value))
    if not candidates:
        return '', set()
    candidates.sort(reverse=True)
    primary_key = candidates[0][2]
    primary = candidates[0][3]
    used_keys = {primary_key}
    supplements = []
    for _, _, key, value in candidates[1:]:
        if value and value not in primary and len(value) <= 400:
            supplements.append(f'{key}：{value}')
            used_keys.add(key)
    return '\n'.join([primary, *supplements]), used_keys


def normalize_record(row):
    """把任意表格行映射为语义清洗输入；保留被拆开的正文与元数据。"""
    row = {
        _as_text(key): _as_text(value)
        for key, value in row.items()
        if key is not None and _as_text(key)
    }
    title = _first_alias(row, COLUMN_ALIASES['标题'])
    link = _first_alias(row, COLUMN_ALIASES['链接'])
    publish_time = _first_alias(row, COLUMN_ALIASES['发布时间'])
    text, used_text_keys = _best_text_value(row)
    region = _first_alias(row, ['地区', '区域', 'region', 'area'])
    deadline = _first_alias(row, ['截止时间', '投标截至时间', '投标截止时间', '报名截止时间', 'deadline', 'end_time'])
    number = _first_alias(row, ['编号', '招标编号', '采购编号', '项目编号', 'code', 'number'])

    # 八爪鱼模板常把公告拆成多个结构化列。没有正文列时，将整行重组成
    # “列名：值”文本，让同一套标签规则继续工作。
    metadata_headers = {
        _normalize_header(alias)
        for aliases in COLUMN_ALIASES.values()
        for alias in aliases
    }
    metadata_headers.update({_normalize_header(alias) for alias in ['编号', '招标编号', '采购编号', '项目编号']})
    supplements = [
        f'{key}：{value}'
        for key, value in row.items()
        if value and key not in used_text_keys and (_normalize_header(key) not in metadata_headers or key in {'文本1', '文本2', '文本3'})
    ]
    if not text:
        text = '\n'.join(supplements)
    elif supplements:
        text = '\n'.join([text, *supplements])

    return {
        '标题': title,
        '链接': link,
        '发布时间': publish_time,
        '文本': text,
        '地区': region,
        '截止时间': deadline,
        '编号': number,
    }


def _read_json(path):
    with path.open('r', encoding='utf-8-sig') as handle:
        payload = json.load(handle)
    if isinstance(payload, list):
        rows = payload
    elif isinstance(payload, dict):
        rows = None
        for key in ('data', 'rows', 'records', 'result', 'items'):
            if isinstance(payload.get(key), list):
                rows = payload[key]
                break
        if rows is None:
            rows = [payload]
    else:
        raise ValueError(f'{path.name}: JSON 顶层必须是对象或数组')
    return [row for row in rows if isinstance(row, dict)]


def _read_csv(path):
    last_error = None
    for encoding in ('utf-8-sig', 'utf-8', 'gb18030'):
        try:
            with path.open('r', encoding=encoding, newline='') as handle:
                sample = handle.read(4096)
                handle.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=',\t;')
                except csv.Error:
                    dialect = csv.excel
                reader = csv.DictReader(handle, dialect=dialect)
                if not reader.fieldnames:
                    raise ValueError(f'{path.name}: 未找到表头')
                return list(reader)
        except UnicodeDecodeError as exc:
            last_error = exc
    raise ValueError(f'{path.name}: 无法识别编码（已尝试 UTF-8/GB18030）') from last_error


def _unique_headers(values):
    headers = []
    counts = {}
    for index, value in enumerate(values, 1):
        base = _as_text(value) or f'未命名列{index}'
        counts[base] = counts.get(base, 0) + 1
        headers.append(base if counts[base] == 1 else f'{base}_{counts[base]}')
    return headers


def _read_xlsx(path, sheet_name=None):
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise RuntimeError(
            '读取 XLSX/XLSM 需要 openpyxl；请运行 '
            '`python -m pip install -r requirements.txt`'
        ) from exc

    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name:
            if sheet_name not in workbook.sheetnames:
                raise ValueError(
                    f'{path.name}: 工作表 {sheet_name!r} 不存在；可选值: {workbook.sheetnames}'
                )
            worksheet = workbook[sheet_name]
        else:
            worksheet = workbook.active
        values = worksheet.iter_rows(values_only=True)
        try:
            headers = _unique_headers(next(values))
        except StopIteration:
            return []
        return [dict(zip(headers, row)) for row in values]
    finally:
        workbook.close()


def _read_xls(path, sheet_name=None):
    try:
        import xlrd
    except ImportError as exc:
        raise RuntimeError(
            '读取旧版 XLS 需要 xlrd；请运行 '
            '`python -m pip install -r requirements.txt`'
        ) from exc

    workbook = xlrd.open_workbook(path)
    worksheet = (
        workbook.sheet_by_name(sheet_name)
        if sheet_name
        else workbook.sheet_by_index(0)
    )
    if worksheet.nrows == 0:
        return []
    headers = _unique_headers(worksheet.row_values(0))
    rows = []
    for row_index in range(1, worksheet.nrows):
        values = []
        for cell in worksheet.row(row_index):
            value = cell.value
            if cell.ctype == xlrd.XL_CELL_DATE:
                value = xlrd.xldate_as_datetime(value, workbook.datemode)
            values.append(value)
        rows.append(dict(zip(headers, values)))
    return rows


def load_records(input_paths, sheet_name=None):
    """读取一个或多个输入文件并统一成标准记录。"""
    records = []
    source_counts = {}
    for input_path in input_paths:
        path = Path(input_path).expanduser().resolve()
        if not path.is_file():
            raise FileNotFoundError(f'输入文件不存在: {path}')
        suffix = path.suffix.lower()
        if suffix == '.json':
            rows = _read_json(path)
        elif suffix in ('.csv', '.tsv'):
            rows = _read_csv(path)
        elif suffix in ('.xlsx', '.xlsm'):
            rows = _read_xlsx(path, sheet_name)
        elif suffix == '.xls':
            rows = _read_xls(path, sheet_name)
        else:
            raise ValueError(
                f'{path.name}: 不支持 {suffix or "无扩展名"}；支持 JSON/CSV/TSV/XLSX/XLSM/XLS'
            )

        normalized = [normalize_record(row) for row in rows]
        normalized = [row for row in normalized if any(row.values())]
        records.extend(normalized)
        source_counts[str(path)] = len(normalized)
    return records, source_counts


def deduplicate_records(records):
    """优先按来源链接去重；无链接时按标题、发布时间和正文前缀去重。"""
    seen = set()
    unique = []
    for record in records:
        link = record.get('链接', '').strip().lower()
        if link:
            key = ('link', link)
        else:
            key = (
                'content',
                record.get('标题', '').strip(),
                record.get('发布时间', '').strip(),
                record.get('文本', '')[:500].strip(),
            )
        if key not in seen:
            seen.add(key)
            unique.append(record)
    return unique


def clean_records(records, verbose=True):
    results = []
    for index, record in enumerate(records, 1):
        cleaned = extract_procurement_text(record)
        results.append(cleaned)
        if verbose:
            non_empty = sum(1 for value in cleaned.values() if value)
            project = cleaned.get('项目名称') or cleaned.get('原标题') or '(无标题)'
            print(f'[{index:>4}/{len(records)}] {non_empty:>2}/29 | {project[:50]}')
    return results


def write_results(results, output_dir, output_name):
    output_dir = Path(output_dir).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    safe_name = re.sub(r'[<>:"/\\|?*]+', '_', output_name).strip('. ') or '招标公告-清洗结果'
    output_csv = output_dir / f'{safe_name}.csv'
    output_json = output_dir / f'{safe_name}.json'

    with output_json.open('w', encoding='utf-8') as handle:
        json.dump(results, handle, ensure_ascii=False, indent=2)
    with output_csv.open('w', encoding='utf-8-sig', newline='') as handle:
        writer = csv.DictWriter(handle, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows(results)
    return output_csv, output_json


def build_parser():
    parser = argparse.ArgumentParser(
        description='将招投标公告 JSON/CSV/Excel 清洗为 29 字段 CSV 和 JSON。'
    )
    parser.add_argument(
        '-i', '--input', nargs='+', required=True, metavar='FILE',
        help='一个或多个 JSON/CSV/TSV/XLSX/XLSM/XLS 文件',
    )
    parser.add_argument(
        '-o', '--output-dir',
        help='输出目录；默认写入第一个输入文件所在目录',
    )
    parser.add_argument(
        '--name', help='输出文件名（不含扩展名）',
    )
    parser.add_argument(
        '--sheet', help='Excel 工作表名；默认读取活动工作表',
    )
    parser.add_argument(
        '--keep-duplicates', action='store_true',
        help='保留重复公告；默认按链接或内容去重',
    )
    parser.add_argument(
        '--quiet', action='store_true', help='不显示逐条处理进度',
    )
    return parser


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        records, source_counts = load_records(args.input, args.sheet)
        raw_count = len(records)
        if not args.keep_duplicates:
            records = deduplicate_records(records)
        if not records:
            raise ValueError('输入文件中没有可清洗的数据行')

        results = clean_records(records, verbose=not args.quiet)
        first_input = Path(args.input[0]).expanduser().resolve()
        output_dir = args.output_dir or first_input.parent
        default_name = (
            f'{first_input.stem}-清洗结果'
            if len(args.input) == 1
            else '招标公告-合并清洗结果'
        )
        output_csv, output_json = write_results(
            results, output_dir, args.name or default_name
        )

        field_stats = {field: sum(bool(row.get(field)) for row in results) for field in FIELDS}
        total_fields = sum(field_stats.values())
        max_fields = len(results) * len(FIELDS)
        print('\n处理完成')
        for source, count in source_counts.items():
            print(f'  输入: {source} ({count} 条)')
        print(f'  去重: {raw_count} -> {len(results)} 条')
        print(f'  字段覆盖: {total_fields}/{max_fields} ({total_fields * 100 / max_fields:.2f}%)')
        print(f'  CSV: {output_csv}')
        print(f'  JSON: {output_json}')
        return 0
    except (OSError, ValueError, RuntimeError, json.JSONDecodeError, csv.Error) as exc:
        parser.exit(2, f'错误: {exc}\n')


if __name__ == '__main__':
    main()
