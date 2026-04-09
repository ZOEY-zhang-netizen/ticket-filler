"""Policy parser for Douyin ticket sheet generation."""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Iterable

from extractor import extract_text

PRODUCT_META = {
    '儿童特惠票': {
        'fee_desc': '儿童票1张',
        'crowd': '1名1M＜身高≤1.4M或3周岁≤年龄≤11周岁的儿童（需至少1名购票成人陪同）',
        'age_limit': '1M<身高≤1.4M或3岁≤年龄≤11岁',
        'is_afternoon': False,
        'entry': '刷身份证入园',
    },
    '双人特惠票': {
        'fee_desc': '双人票1张',
        'crowd': '2名成人（需同时入园）',
        'age_limit': '',
        'is_afternoon': False,
        'entry': '刷身份证入园',
    },
    '家庭特惠票': {
        'fee_desc': '家庭票1张',
        'crowd': '2名成人+1名1M＜身高≤1.4M或3周岁≤年龄≤11周岁的儿童（需同时入园）',
        'age_limit': '',
        'is_afternoon': False,
        'entry': '刷身份证入园',
    },
    '下午场特惠票': {
        'fee_desc': '下午场票1张',
        'crowd': '1名成人',
        'age_limit': '',
        'is_afternoon': True,
        'entry': '刷身份证15:00（含）以后入园',
    },
    '合家欢套票': {
        'fee_desc': '合家欢套票1张',
        'crowd': '2名成人+2名优待人群（1M＜身高≤1.4M或3周岁≤年龄≤11周岁的儿童或55周岁及以上长者，需同时入园）',
        'age_limit': '',
        'is_afternoon': False,
        'entry': '刷身份证入园',
    },
    '马年生肖票': {
        'fee_desc': '生肖票1张',
        'crowd': '1名生肖游客（以身份证标注生肖为准）',
        'age_limit': '',
        'is_afternoon': False,
        'entry': '刷身份证入园',
    },
    '成人特惠票': {
        'fee_desc': '成人票1张',
        'crowd': '1名成人',
        'age_limit': '',
        'is_afternoon': False,
        'entry': '刷身份证入园',
    },
    '闺蜜双人票': {
        'fee_desc': '双人票1张',
        'crowd': '2名女性（需同时入园）',
        'age_limit': '',
        'is_afternoon': False,
        'entry': '刷身份证入园',
    },
}

HEADING_RE = re.compile(r'.{0,20}(政策|活动政策|活动)$')
CHANNEL_KEYWORDS = ('抖音', '抖团')
EXCLUDE_CHANNEL_KEYWORDS = ('美团', '携程', '山姆', '京东', '同程', '飞猪')
EXCLUDE_SECTION_KEYWORDS = ('酒店', '一卡通', '党群', '临港一卡通')
IMPLICIT_DOUYIN_SECTION_KEYWORDS = ('全渠道', '线上渠道', '直播间', '平台活动政策', 'POI')
DETAIL_KEYWORDS = (
    '销售时间', '使用时间', '适用人群', '使用人群', '购买规则', '退改规则', '退改说明',
    '退款规则', '退票规则', '销售渠道', '销售渠道及库存', '渠道及库存',
    '活动时间', '预定规则', '加价规则', '划拨价', '销售说明',
)
NON_PRODUCT_PREFIXES = (
    '销售时间', '使用时间', '适用人群', '使用人群', '购买规则', '退改规则', '退改说明',
    '退款规则', '退票规则', '销售渠道', '销售渠道及库存', '渠道及库存',
    '门市价', '对标售价', '对标当前售价', '现售价', '活动价', '销售价',
    '结算价', '门市价折扣', '备注', '产品', '票种', '门票类型', '日期划分',
    '渠道类型', '说明', '活动时间', '预定规则', '加价规则', '划拨价', '销售说明',
)
NON_TICKET_KEYWORDS = (
    '房', '酒店', '客房', '早餐', '剧场', '演出', '餐厅', '餐饮',
    '代金券', '优惠券', '获客卡', '房券', '票根', '住宿', '赛事包'
)
PRODUCT_NAME_RE = re.compile(r'^(.{1,40}?(?:门票|套票|联票|免费票|特惠票|秒杀票|双人票|家庭票|亲子票|下午场票|双次票|票))')
DATE_RANGE_RE = re.compile(
    r'(?P<start>(?:(?:20)?\d{2})年\s*\d{1,2}月\d{1,2}日|\d{1,2}月\d{1,2}日)'
    r'\s*[－\-~～至到]\s*'
    r'(?P<end>(?:(?:20)?\d{2})年\s*\d{1,2}月\d{1,2}日|\d{1,2}月\d{1,2}日)'
)
DATE_TOKEN_RE = re.compile(r'((?:(?:20)?\d{2})年)?\s*(\d{1,2})月(\d{1,2})日')
VALUE_TOKEN_RE = re.compile(
    r'(?:平日|高峰日|春节|非春节|通用|特别常规日|常规日|工作日|周末|节假日|非节假日|成人|儿童|优待|标准)[：:]\s*[-免费\d./]+'
    r'|(?:免费|\d+(?:\.\d+)?)元?[：:].*'
    r'|免费|-|\d+(?:\.\d+)?(?:/\d+(?:\.\d+)?)?'
)
USE_MANUAL_REVIEW_HINTS = ('①', '②', '③', '④', '根据', '开放时间', '动态释放', '视预订情况', '首次入园', '第二次入园')
HEADER_PATTERNS = [
    ('discount', r'门市价折扣'),
    ('sale_price', r'25年销售价(?:（[^）]+）)?|销售价|活动价'),
    ('settlement_price', r'25年(?:渠道)?结算价|结算价'),
    ('list_price', r'25年门市价(?:（[^）]+）)?|门市价'),
    ('compare_price', r'对标当前售价|25年对标售价|24年对标价格|24年同期|对标售价|现售价|12月份政策|25年对标价'),
    ('remarks', r'备注'),
    ('product', r'产品'),
]
DEFAULT_HEADER_SEMANTICS = ['product', 'list_price', 'compare_price', 'sale_price', 'settlement_price', 'discount', 'remarks']


def normalize_text(text: str) -> str:
    text = text.replace('\r', '\n')
    text = re.sub(r'\u3000', ' ', text)
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r' *\n *', '\n', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


def _expand_year(raw_year: str | None) -> int | None:
    if not raw_year:
        return None
    year = int(raw_year[:-1] if raw_year.endswith('年') else raw_year)
    return 2000 + year if year < 100 else year


def infer_year(text: str, source_name: str | None = None) -> int:
    for scope in [source_name or '', '\n'.join(text.splitlines()[:50]), text[:3000]]:
        for pat in [r'((?:20)?\d{2})年全年', r'((?:20)?\d{2})年[^\n]{0,20}(?:票价|门票|价格政策|活动政策)', r'((?:20)?\d{2})年\d{1,2}月\d{1,2}日']:
            m = re.search(pat, scope)
            if m:
                y = _expand_year(m.group(1))
                if y:
                    return y
    return date.today().year


def parse_mmdd(mmdd: str, year: int) -> str | None:
    m = DATE_TOKEN_RE.search(mmdd)
    if not m:
        return None
    explicit_year = _expand_year(m.group(1))
    month, day = int(m.group(2)), int(m.group(3))
    actual_year = explicit_year or year
    return f'{actual_year:04d}-{month:02d}-{day:02d}'


def split_sections(text: str) -> list[tuple[str, str]]:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    sections: list[tuple[str, list[str]]] = []
    current_title = '未命名部分'
    current_lines: list[str] = []
    for line in lines:
        if HEADING_RE.search(line):
            if current_lines:
                sections.append((current_title, current_lines))
            current_title = line
            current_lines = []
        else:
            current_lines.append(line)
    if current_lines:
        sections.append((current_title, current_lines))
    return [(title, '\n'.join(body)) for title, body in sections]


def contains_douyin_channel(block: str, section_title: str = '') -> bool:
    combined = f'{section_title}\n{block}'
    if any(k in combined for k in CHANNEL_KEYWORDS):
        return True
    if any(k in section_title for k in IMPLICIT_DOUYIN_SECTION_KEYWORDS):
        return True
    if any(k in combined for k in EXCLUDE_CHANNEL_KEYWORDS):
        return False
    return True


def should_skip_section(title: str) -> bool:
    return any(k in title for k in EXCLUDE_SECTION_KEYWORDS)


def _strip_line_prefix(line: str) -> str:
    line = line.strip()
    line = re.sub(r'^[*＊•\-]\s*', '', line)
    line = re.sub(r'^[0-9一二三四五六七八九十]+[、．)）]\s*', '', line)
    return line.strip()


def _clean_product_name(name: str) -> str:
    name = _strip_line_prefix(name)
    name = re.sub(r'^(?:[【\[].*?[】\]]\s*)+', '', name)
    name = re.sub(r'\s+', '', name)
    for keyword in DETAIL_KEYWORDS:
        if keyword in name:
            name = name.split(keyword, 1)[0]
    m = PRODUCT_NAME_RE.match(name)
    if m:
        name = m.group(1)
    return name.strip('：:；;，, ')


def _extract_candidate_name(line: str) -> str | None:
    line = _strip_line_prefix(line)
    if not line or any(line.startswith(prefix) for prefix in NON_PRODUCT_PREFIXES):
        return None
    head = line.split('\t', 1)[0].strip()
    name = _clean_product_name(head)
    colon_positions = [pos for pos in (head.find('：'), head.find(':')) if pos >= 0]
    if colon_positions and min(colon_positions) < len(name):
        return None
    if not name or len(name) > 40:
        return None
    return name


def _is_ticket_product(name: str, block: str) -> bool:
    if not name or name in {'产品', '备注', '说明'}:
        return False
    if not any(k in name for k in ('票', '套票', '联票')):
        return False
    if any(k in name for k in NON_TICKET_KEYWORDS):
        return False
    return True


def find_product_blocks(section_body: str) -> Iterable[tuple[str, str, int]]:
    matches: list[tuple[int, str]] = []
    cursor = 0
    for raw_line in section_body.splitlines(True):
        line = raw_line.strip()
        if line:
            name = _extract_candidate_name(line)
            if name and _is_ticket_product(name, line):
                matches.append((cursor + raw_line.find(line), name))
        cursor += len(raw_line)
    filtered: list[tuple[int, str]] = []
    seen: set[int] = set()
    for idx, (pos, name) in enumerate(matches):
        if pos in seen:
            continue
        if idx + 1 < len(matches) and matches[idx + 1][1] == name and matches[idx + 1][0] - pos < 300:
            continue
        seen.add(pos)
        filtered.append((pos, name))
    for idx, (start, name) in enumerate(filtered):
        end = filtered[idx + 1][0] if idx + 1 < len(filtered) else len(section_body)
        yield name, section_body[start:end].strip(), start


def section_label(title: str) -> str:
    if '清明' in title:
        return '清明节预售'
    m = re.search(r'(\d[\d\-]*月)', title)
    if m:
        return f'{m.group(1)}全渠道'
    return title.strip()


def split_compound_block(product_name: str, block: str) -> list[str]:
    lines = [line.strip() for line in block.splitlines() if line.strip()]
    sale_indexes = [i for i, raw in enumerate(lines) if _strip_line_prefix(raw).startswith('销售时间')]
    if len(sale_indexes) <= 1:
        return [block]
    chunks: list[str] = []
    cursor = 1 if lines and product_name in lines[0] and '销售时间' in lines[0] else 0
    while cursor < len(lines):
        next_sale = next((i for i in range(cursor, len(lines)) if _strip_line_prefix(lines[i]).startswith('销售时间')), None)
        if next_sale is None:
            break
        end = next((i for i in range(next_sale, len(lines)) if any(tok in _strip_line_prefix(lines[i]) for tok in ('退改规则', '退改说明', '退款规则', '退票规则'))), len(lines) - 1)
        chunk_lines = lines[cursor:end + 1]
        if chunk_lines and chunk_lines[0] != product_name:
            chunk_lines = [product_name] + chunk_lines
        chunks.append('\n'.join(chunk_lines))
        cursor = end + 1
    return chunks or [block]


def _strip_storage_prices(text: str) -> str:
    text = re.sub(r'\d+(?:\.\d+)?元?\s*[（(]储备[)）]', ' ', text)
    text = re.sub(r'\d+(?:\.\d+)?元为储备[^；。\n]*', ' ', text)
    text = re.sub(r'\n\d+(?:\.\d+)?\s*\n[（(]?储备[)）]?', '\n', text)
    return text


def _extract_price_area(block: str) -> str:
    lines = [line.strip() for line in block.splitlines() if line.strip()]
    structured: list[str] = []
    for raw in lines[1:]:
        display = raw.strip()
        head = re.sub(r'^\d+[、)）]\s*', '', display)
        if any(head.startswith(k) for k in DETAIL_KEYWORDS):
            break
        if any(k in display for k in DETAIL_KEYWORDS):
            if structured:
                break
            continue
        stripped = _strip_storage_prices(display).strip()
        if stripped:
            structured.append(stripped)
    if structured:
        return '\n'.join(structured)
    cut_positions = [block.find(keyword) for keyword in DETAIL_KEYWORDS if keyword in block]
    cut_positions = [pos for pos in cut_positions if pos >= 0]
    cut = min(cut_positions) if cut_positions else len(block)
    return _strip_storage_prices(block[:cut]).strip()


def _find_header_text(prefix_text: str) -> str:
    lines = [line.strip() for line in prefix_text.splitlines() if line.strip()]
    recent = lines[-12:]
    for line in reversed(recent):
        if '产品' in line and any(k in line for k in ('销售价', '活动价', '结算价', '门市价', '现售价')):
            return line.replace(' ', '')
    product_indexes = [idx for idx, line in enumerate(recent) if line == '产品' or line.startswith('产品')]
    if product_indexes:
        joined = ''.join(recent[product_indexes[-1]:]).replace(' ', '')
        if any(k in joined for k in ('销售价', '活动价', '结算价', '门市价', '现售价')):
            return joined
    return ''


def _extract_header_semantics(prefix_text: str) -> list[str]:
    header_text = _find_header_text(prefix_text)
    if not header_text:
        return DEFAULT_HEADER_SEMANTICS[:]
    spans: list[tuple[int, int, str]] = []
    occupied: list[tuple[int, int]] = []
    for semantic, pattern in HEADER_PATTERNS:
        for match in re.finditer(pattern, header_text):
            start, end = match.span()
            if any(not (end <= s or start >= e) for s, e in occupied):
                continue
            spans.append((start, end, semantic))
            occupied.append((start, end))
    spans.sort(key=lambda x: x[0])
    semantics = [semantic for _, _, semantic in spans]
    return semantics if semantics and 'sale_price' in semantics else DEFAULT_HEADER_SEMANTICS[:]


def _is_value_like_line(line: str) -> bool:
    line = line.strip()
    if not line or any(line.startswith(prefix) for prefix in DETAIL_KEYWORDS):
        return False
    if re.match(r'^(?:免费|\d+(?:\.\d+)?)元?[：:]', line):
        return True
    if re.match(r'^(?:平日|高峰日|春节|非春节|通用|特别常规日|常规日|工作日|周末|节假日|非节假日|成人|儿童|优待|标准)[：:]', line):
        return True
    return bool(re.fullmatch(r'免费|-|\d+(?:\.\d+)?(?:/\d+(?:\.\d+)?)?|[-免费\d./：: ]+', line))


def _value_line_kind(line: str) -> str:
    if _value_line_label(line):
        return 'day_labeled'
    if re.match(r'^(?:免费|\d+(?:\.\d+)?)元?[：:]', line):
        return 'labeled_text'
    return 'plain'


def _value_line_label(line: str) -> str | None:
    m = re.match(r'^(平日|高峰日|春节|非春节|通用|特别常规日|常规日|工作日|周末|节假日|非节假日|成人|儿童|优待|标准)[：:]', line)
    return m.group(1) if m else None


def _merge_value_lines(lines: list[str]) -> list[str]:
    merged: list[str] = []
    buffer: list[str] = []
    buffer_kind: str | None = None
    for raw in lines:
        line = _strip_storage_prices(raw).strip()
        if not line:
            continue
        kind = _value_line_kind(line) if _is_value_like_line(line) else None
        if kind in {'day_labeled', 'labeled_text'}:
            if buffer and buffer_kind == kind:
                buffer.append(line)
            else:
                if buffer:
                    merged.append(' '.join(buffer))
                buffer = [line]
                buffer_kind = kind
            continue
        if buffer:
            merged.append(' '.join(buffer))
            buffer = []
            buffer_kind = None
        merged.append(line)
    if buffer:
        merged.append(' '.join(buffer))
    return merged


def _extract_value_tokens(price_area: str, product_name: str) -> list[str]:
    lines = [line.strip() for line in price_area.splitlines() if line.strip()]
    standalone: list[str] = []
    if len(lines) > 1:
        for line in lines:
            if _clean_product_name(line) == product_name:
                continue
            if _is_value_like_line(line):
                standalone.append(line)
        if standalone:
            return _merge_value_lines(standalone)
    first_line = _strip_storage_prices(lines[0]) if lines else ''
    tail = first_line.split(product_name, 1)[1] if product_name and product_name in first_line else first_line
    return _merge_value_lines([token.strip() for token in VALUE_TOKEN_RE.findall(tail) if token.strip()])


def _token_to_number(value: str | None) -> int | float | None:
    if not value:
        return None
    if '免费' in value and not re.search(r'\d+(?:\.\d+)?', value):
        return 0
    match = re.search(r'\d+(?:\.\d+)?', value)
    if not match:
        return None
    number = float(match.group(0))
    return int(number) if number.is_integer() else number


def _token_to_numbers(value: str | None) -> list[int | float]:
    if not value:
        return []
    if '免费' in value and not re.search(r'\d+(?:\.\d+)?', value):
        return [0]
    numbers: list[int | float] = []
    for m in re.finditer(r'\d+(?:\.\d+)?', value):
        n = float(m.group(0))
        numbers.append(int(n) if n.is_integer() else n)
    return numbers


def _map_price_values(header_semantics: list[str], tokens: list[str]) -> dict[str, list[str]]:
    semantics = [s for s in header_semantics if s not in ('product', 'remarks')]
    if len(tokens) == 3 and any('：' in tokens[-1] or ':' in tokens[-1] for _ in [0]):
        semantics = ['sale_price', 'settlement_price', 'discount']
    elif len(tokens) == 4 and (('：' in tokens[0] or ':' in tokens[0]) and ('：' in tokens[-1] or ':' in tokens[-1])):
        semantics = ['list_price', 'sale_price', 'settlement_price', 'discount']
    else:
        while len(semantics) > len(tokens):
            for key in ('compare_price', 'settlement_price', 'list_price'):
                idxs = [i for i, semantic in enumerate(semantics) if semantic == key]
                if idxs:
                    semantics.pop(idxs[0] if key != 'settlement_price' else idxs[-1])
                    break
            else:
                break
    mapped: dict[str, list[str]] = {}
    for semantic, value in zip(semantics, tokens):
        mapped.setdefault(semantic, []).append(value)
    return mapped


def extract_list_price(block: str, header_semantics: list[str] | None = None, product_name: str = '') -> int | float | None:
    header_semantics = header_semantics or DEFAULT_HEADER_SEMANTICS
    price_area = _extract_price_area(block)
    tokens = _extract_value_tokens(price_area, product_name)
    mapped = _map_price_values(header_semantics, tokens)
    if mapped.get('list_price'):
        nums = _token_to_numbers(mapped['list_price'][0])
        if nums:
            return max(nums)
    peak = re.search(r'高峰日[：:]\s*(\d+(?:\.\d+)?)', price_area)
    if peak:
        return _token_to_number(peak.group(1))
    nums: list[int | float] = []
    for token in tokens:
        nums.extend(_token_to_numbers(token))
    return max(nums) if nums else None


def extract_sale_price_variants(block: str, header_semantics: list[str] | None = None, product_name: str = '') -> list[int | float]:
    header_semantics = header_semantics or DEFAULT_HEADER_SEMANTICS
    price_area = _extract_price_area(block)
    tokens = _extract_value_tokens(price_area, product_name)
    if any('核销' in token for token in tokens):
        plain_numeric = [token for token in tokens if not ('：' in token or ':' in token) and _token_to_numbers(token)]
        if plain_numeric:
            return _token_to_numbers(plain_numeric[0])
    mapped = _map_price_values(header_semantics, tokens)
    if 'sale_price' in mapped:
        return _token_to_numbers(mapped['sale_price'][0])
    nums: list[int | float] = []
    for token in tokens:
        nums.extend(_token_to_numbers(token))
    return [nums[0]] if nums else []


def extract_sale_price(block: str, header_semantics: list[str] | None = None, product_name: str = '') -> int | float | None:
    variants = extract_sale_price_variants(block, header_semantics, product_name)
    return variants[0] if variants else None


def extract_refund(block: str, product_name: str) -> str:
    for raw in block.splitlines():
        cleaned = _strip_line_prefix(raw)
        m = re.search(r'(?:退改规则|退改说明|退款规则|退票规则)[：:]?\s*(.*)', cleaned)
        if m:
            val = m.group(1).strip().rstrip('；;')
            if val:
                return val if val.endswith('。') else val + '。'
    if '双人' in product_name:
        return '已使用不可退，未使用可退，单张双人特惠票不支持部分退款。'
    if '家庭' in product_name or '合家欢' in product_name:
        return '已使用不可退，未使用可退，单张家庭特惠票不支持部分退款。'
    return '已使用不可退，未使用可退。'


def extract_crowd(block: str) -> str | None:
    for raw in block.splitlines():
        cleaned = _strip_line_prefix(raw)
        if cleaned.startswith('适用人群') or cleaned.startswith('使用人群'):
            return re.sub(r'^(?:适用人群|使用人群)[：:]?', '', cleaned).strip().rstrip('；;。')
    return None


def extract_limit(block: str, keyword: str, channel: str, default: int | None = 1) -> int | None:
    limit_phrase = r'(?:限购|仅限购|仅限购买)'
    patterns: list[str] = []
    if channel == 'tuan':
        if keyword == '身份证':
            patterns.extend([
                rf'活动期间同一身份证(?:号)?及手机号(?:号)?{limit_phrase}(\d+)张',
                rf'活动期间同一身份证(?:号)?{limit_phrase}(\d+)张',
            ])
        else:
            patterns.extend([
                rf'活动期间同一身份证(?:号)?及手机号(?:号)?{limit_phrase}(\d+)张',
                rf'活动期间同一手机号(?:号)?{limit_phrase}(\d+)张',
            ])
    else:
        if keyword == '身份证':
            patterns.extend([
                rf'游玩日当天同一身份证(?:号)?及手机号(?:号)?{limit_phrase}(\d+)张',
                rf'游玩日当天同一身份证(?:号)?{limit_phrase}(\d+)张',
                rf'活动期间同一身份证(?:号)?{limit_phrase}(\d+)张',
            ])
        else:
            patterns.extend([
                rf'游玩日当天同一身份证(?:号)?及手机号(?:号)?{limit_phrase}(\d+)张',
                rf'游玩日当天同一手机号(?:号)?{limit_phrase}(\d+)张',
                rf'活动期间同一手机号(?:号)?{limit_phrase}(\d+)张',
            ])
    for pattern in patterns:
        m = re.search(pattern, block)
        if m:
            return int(m.group(1))
    return default


def extract_purchase_rules(block: str) -> tuple[str | None, str | None]:
    lines = [line.strip() for line in block.splitlines() if line.strip()]
    collecting = False
    gathered: list[str] = []
    inline_rule = None
    for raw in lines:
        cleaned = _strip_line_prefix(raw)
        if cleaned.startswith('购买规则'):
            collecting = True
            after = re.sub(r'^购买规则[：:]?', '', cleaned).strip()
            if after:
                inline_rule = after.rstrip('；;')
            continue
        if not collecting:
            continue
        if any(tok in cleaned for tok in ('退改规则', '退改说明', '退款规则', '退票规则')):
            break
        if re.match(r'^\d+[、.．）)]', raw):
            if any(tok in cleaned for tok in ('线下渠道', '储备政策')):
                continue
        gathered.append(cleaned.rstrip('；;'))
    semi_rule = None
    tuan_rule = None
    generic_rule = inline_rule
    for line in gathered:
        compact = line.replace(' ', '')
        if compact.startswith('POI：') or compact.startswith('POI:'):
            semi_rule = line
        elif compact.startswith('其他渠道：') or compact.startswith('其他渠道:'):
            tuan_rule = line
        elif compact.startswith('抖音') or compact.startswith('直播间') or '期票' in line:
            tuan_rule = line
        elif generic_rule is None:
            generic_rule = line
    if semi_rule is None:
        semi_rule = generic_rule
    if tuan_rule is None:
        tuan_rule = generic_rule
    return semi_rule, tuan_rule


def _apply_cross_year(start_iso: str | None, end_iso: str | None, start_token: str, end_token: str) -> tuple[str | None, str | None]:
    if not start_iso or not end_iso:
        return start_iso, end_iso
    start_year, start_month, _ = map(int, start_iso.split('-'))
    end_year, end_month, end_day = map(int, end_iso.split('-'))
    if '年' not in end_token and end_month < start_month:
        end_year = start_year + 1
        end_iso = f'{end_year:04d}-{end_month:02d}-{end_day:02d}'
    return start_iso, end_iso


def extract_date_range(block: str, label: str, year: int) -> tuple[str | None, str | None, bool, str | None]:
    lines = [line.strip() for line in block.splitlines() if line.strip()]
    line = None
    for raw in lines:
        cleaned = _strip_line_prefix(raw)
        if cleaned.startswith(label):
            line = re.sub(rf'^{label}[：:]?', '', cleaned).strip()
            break
    if line is None:
        return None, None, False, None
    ranges = list(DATE_RANGE_RE.finditer(line))
    manual_review = label == '使用时间' and (len(ranges) > 1 or any(h in line for h in USE_MANUAL_REVIEW_HINTS))
    if ranges:
        first = ranges[0]
        start_token = first.group('start')
        end_token = first.group('end')
        start_iso = parse_mmdd(start_token, year)
        start_year = int(start_iso[:4]) if start_iso else year
        end_iso = parse_mmdd(end_token, start_year)
        start_iso, end_iso = _apply_cross_year(start_iso, end_iso, start_token, end_token)
        return start_iso, end_iso, manual_review, line
    single = DATE_TOKEN_RE.search(line)
    if single:
        d = parse_mmdd(single.group(0), year)
        return d, d, manual_review, line
    return None, None, manual_review, line


def extract_use_price_variants(use_line: str | None, year: int) -> list[dict[str, object]]:
    if not use_line:
        return []
    variants: list[dict[str, object]] = []
    pattern = re.compile(r'(?:[①②③④⑤⑥]\s*)?(\d+(?:\.\d+)?)元[：:]\s*([^①②③④⑤⑥]+?)(?=(?:[①②③④⑤⑥]\s*)?\d+(?:\.\d+)?元[：:]|$)')
    for m in pattern.finditer(use_line):
        num = float(m.group(1))
        price = int(num) if num.is_integer() else num
        segment = m.group(2).strip().rstrip('；;')
        ranges = list(DATE_RANGE_RE.finditer(segment))
        if not ranges:
            continue
        first = ranges[0]
        start_token = first.group('start')
        end_token = first.group('end')
        start_iso = parse_mmdd(start_token, year)
        start_year = int(start_iso[:4]) if start_iso else year
        end_iso = parse_mmdd(end_token, start_year)
        start_iso, end_iso = _apply_cross_year(start_iso, end_iso, start_token, end_token)
        variants.append({
            'price': price,
            'use_start': start_iso,
            'use_end': end_iso,
            'use_manual_review': False,
            'use_rule_raw': f'{price}元：{segment}',
        })
    return variants




def trim_block_after_refund(block: str) -> str:
    lines = [line.strip() for line in block.splitlines() if line.strip()]
    for idx, raw in enumerate(lines):
        cleaned = _strip_line_prefix(raw)
        if any(tok in cleaned for tok in ('退改规则', '退改说明', '退款规则', '退票规则')):
            return '\n'.join(lines[:idx + 1])
    return block

def _resolve_product_meta(product_name: str, block: str) -> dict:
    if product_name in PRODUCT_META:
        meta = dict(PRODUCT_META[product_name])
    elif '下午场' in product_name:
        meta = dict(PRODUCT_META['下午场特惠票'])
    elif '合家欢' in product_name:
        meta = dict(PRODUCT_META['合家欢套票'])
    elif '家庭' in product_name:
        meta = dict(PRODUCT_META['家庭特惠票'])
    elif '儿童' in product_name:
        meta = dict(PRODUCT_META['儿童特惠票'])
    elif '闺蜜' in product_name:
        meta = dict(PRODUCT_META['闺蜜双人票'])
    elif '双人' in product_name:
        meta = dict(PRODUCT_META['双人特惠票'])
    elif '生肖' in product_name:
        meta = dict(PRODUCT_META['马年生肖票'])
    elif '成人' in product_name:
        meta = dict(PRODUCT_META['成人特惠票'])
    else:
        meta = {'fee_desc': f'{product_name}1张', 'crowd': '1名成人', 'age_limit': '', 'is_afternoon': False, 'entry': '刷身份证入园'}
    if '15:00' in block or '下午场' in product_name:
        meta['is_afternoon'] = True
        meta['entry'] = '刷身份证15:00（含）以后入园'
    if '大学生' in product_name:
        meta['age_limit'] = '仅限全日制专本科、对口高职高等教育学校在校生'
    return meta


def parse_policy_text(text: str, source_name: str | None = None) -> list[dict]:
    text = normalize_text(text)
    year = infer_year(text, source_name=source_name)
    products: list[dict] = []
    seen: set[tuple] = set()
    for title, body in split_sections(text):
        if should_skip_section(title):
            continue
        label = section_label(title)
        for product_name, block, start in find_product_blocks(body):
            if not contains_douyin_channel(block, title):
                continue
            header_semantics = _extract_header_semantics(body[:start])
            parent_list_price = extract_list_price(block, header_semantics, product_name)
            for subblock in split_compound_block(product_name, block):
                subblock = trim_block_after_refund(subblock)
                list_price = extract_list_price(subblock, header_semantics, product_name) or parent_list_price
                sale_price = extract_sale_price(subblock, header_semantics, product_name)
                sale_variants = extract_sale_price_variants(subblock, header_semantics, product_name)
                sale_start, sale_end, _, sale_line = extract_date_range(subblock, '销售时间', year)
                use_start, use_end, use_manual_review, use_line = extract_date_range(subblock, '使用时间', year)
                if not sale_start or not sale_end or not use_start or list_price is None or sale_price is None:
                    continue
                meta = _resolve_product_meta(product_name, subblock)
                crowd = extract_crowd(subblock) or meta['crowd']
                purchase_rule_semi, purchase_rule_tuan = extract_purchase_rules(subblock)
                hotel_face_allowed = '录脸入园' in subblock and not meta.get('is_afternoon')
                use_variants = extract_use_price_variants(use_line, year)
                item_variants: list[dict[str, object]] = []
                if use_variants and sale_variants:
                    sale_set = {float(v) for v in sale_variants}
                    matched = [variant for variant in use_variants if float(variant['price']) in sale_set]
                    if matched:
                        item_variants = matched
                if not item_variants:
                    item_variants = [{
                        'price': sale_price,
                        'use_start': use_start,
                        'use_end': use_end if not meta['is_afternoon'] else None,
                        'use_manual_review': use_manual_review,
                        'use_rule_raw': use_line,
                    }]
                for variant in item_variants:
                    item = {
                        'name': product_name,
                        'section': label,
                        'year': year,
                        'list_price': list_price,
                        'sale_price': variant['price'],
                        'sale_start': sale_start,
                        'sale_end': sale_end,
                        'sale_rule_raw': sale_line,
                        'use_start': variant['use_start'],
                        'use_end': variant['use_end'] if not meta['is_afternoon'] else None,
                        'use_manual_review': bool(variant.get('use_manual_review', False)),
                        'use_rule_raw': variant.get('use_rule_raw') or use_line,
                        'refund': extract_refund(subblock, product_name),
                        'id_limit_tuan': extract_limit(subblock, '身份证', 'tuan', 1),
                        'id_limit_semi': extract_limit(subblock, '身份证', 'semi', 1),
                        'phone_limit_tuan': extract_limit(subblock, '手机号', 'tuan', None),
                        'phone_limit_semi': extract_limit(subblock, '手机号', 'semi', None),
                        'purchase_rule_tuan': purchase_rule_tuan,
                        'purchase_rule_semi': purchase_rule_semi,
                        'hotel_face_allowed': hotel_face_allowed,
                        **meta,
                        'crowd': crowd,
                    }
                    key = (item['section'], item['name'], item['sale_price'], item['sale_start'], item['sale_end'], item['use_start'], item['use_end'])
                    if key in seen:
                        continue
                    seen.add(key)
                    products.append(item)
    return products


def parse_policy(path_or_text: str, **kwargs) -> list[dict]:
    if Path(path_or_text).exists():
        text = extract_text(path_or_text)
        return parse_policy_text(text, source_name=Path(path_or_text).name)
    return parse_policy_text(path_or_text)
