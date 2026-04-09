"""
main.py  ─  抖音门票政策自动化建票工具
────────────────────────────────────────
Usage:
    python main.py <policy.docx> [output.xlsx] [--api-key sk-ant-...]

Example:
    export ANTHROPIC_API_KEY='sk-ant-...'
    python main.py 【门票】上海项目清明预售及4月活动政策申请.docx

    # 或者直接传入 API key:
    python main.py policy.docx --api-key sk-ant-...

输出文件默认放在与 docx 同目录，名称自动加 -filled.xlsx 后缀。
"""

import sys
import json
import argparse
from pathlib import Path

# Allow running from parent directory
sys.path.insert(0, str(Path(__file__).parent))
from policy_parser import parse_policy
from excel_writer import write_excel

TEMPLATE = str(
    Path(__file__).parent.parent
    / '【抖音】上海项目4月门票及酒店价格政策申请建票-0404 (2).xlsx'
)


def main():
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument('docx', nargs='?')
    parser.add_argument('output', nargs='?')
    parser.add_argument('--api-key', default=None)
    args, _ = parser.parse_known_args()

    if not args.docx:
        print(__doc__)
        sys.exit(1)

    docx_path = args.docx
    if not Path(docx_path).exists():
        print(f"Error: file not found → {docx_path}")
        sys.exit(1)

    output_path = args.output or str(
        Path(docx_path).parent / f'{Path(docx_path).stem}-filled.xlsx'
    )

    print(f"[1/3] 正在解析政策文档: {Path(docx_path).name}")
    products = parse_policy(docx_path, api_key=args.api_key)
    print(f"      提取到 {len(products)} 个抖音产品：")
    for p in products:
        print(f"      · [{p.get('section','')}] {p['name']}  "
              f"售价={p['sale_price']}  结算={p.get('settle_price','auto')}")

    print(f"\n[2/3] 正在写入 Excel 模版...")
    write_excel(products, TEMPLATE, output_path)

    print(f"\n[3/3] 完成！输出文件：")
    print(f"      {output_path}")
    print(f"\n      共写入 {len(products) * 2} 行（每个产品 2 个渠道）")
    print(f"      🔴 红色单元格：活动标签（F列）请手动选择；"
          f"下午场使用时间止（M列）请手动填写")

    # Also save parsed JSON for reference
    json_path = output_path.replace('.xlsx', '_parsed.json')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(products, f, ensure_ascii=False, indent=2)
    print(f"\n      解析结果已保存至: {Path(json_path).name}")


if __name__ == '__main__':
    main()
