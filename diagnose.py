# -*- coding: utf-8 -*-
"""
诊断脚本：检查 PDF 空白页删除是否正常工作
使用方法：先转换一个 Excel 文件，然后运行：
  python 诊断空白页.py output/xxx.pdf
"""
import sys, os, re

# 检查 pypdf
print("=" * 50)
print("  PDF 空白页诊断工具")
print("=" * 50)

try:
    from pypdf import PdfReader
    print("\n✅ pypdf 已安装")
except ImportError:
    print("\n❌ pypdf 未安装！这就是空白页删除不生效的原因！")
    print("   请运行: pip install pypdf")
    input("\n按回车退出...")
    sys.exit(1)

# 找 PDF 文件
if len(sys.argv) > 1:
    pdf_path = sys.argv[1]
else:
    # 在 output 目录找第一个 PDF
    output_dir = os.path.join(os.path.dirname(__file__), "output")
    pdfs = []
    if os.path.exists(output_dir):
        for root, dirs, files in os.walk(output_dir):
            for f in files:
                if f.endswith('.pdf'):
                    pdfs.append(os.path.join(root, f))
    if pdfs:
        pdf_path = pdfs[0]
        print(f"\n自动找到: {pdf_path}")
    else:
        print("\n❌ 未找到 PDF 文件")
        print("   请先转换一个 Excel，或指定路径: python 诊断空白页.py xxx.pdf")
        input("\n按回车退出...")
        sys.exit(1)

if not os.path.exists(pdf_path):
    print(f"\n❌ 文件不存在: {pdf_path}")
    input("\n按回车退出...")
    sys.exit(1)

# 分析每一页
reader = PdfReader(pdf_path)
total = len(reader.pages)
print(f"\n文件: {os.path.basename(pdf_path)}")
print(f"总页数: {total}")
print("-" * 50)

for i, page in enumerate(reader.pages):
    text = page.extract_text() or ""
    has_text = bool(re.search(r'[\w\u4e00-\u9fff]', text))
    text_preview = text.strip()[:80].replace('\n', ' ') if text.strip() else "(空)"

    # 检查图片
    has_real_image = False
    xobj_list = []
    try:
        if '/Resources' in page:
            resources = page['/Resources']
            if '/XObject' in resources:
                xobjects = resources['/XObject']
                if xobjects:
                    for key in xobjects:
                        try:
                            xobj = xobjects[key]
                            obj = xobj.get_object() if hasattr(xobj, 'get_object') else xobj
                            subtype = str(obj.get('/Subtype', '?'))
                            xobj_list.append(f"{key}({subtype})")
                            if subtype == '/Image':
                                has_real_image = True
                        except:
                            xobj_list.append(f"{key}(错误)")
    except:
        pass

    is_blank = not has_text and not has_real_image
    status = "🔴 空白页" if is_blank else "🟢 有内容"

    print(f"\n第 {i+1} 页: {status}")
    print(f"  文字: {'有' if has_text else '无'} → {text_preview}")
    print(f"  图片: {'有' if has_real_image else '无'}")
    if xobj_list:
        print(f"  XObject: {', '.join(xobj_list)}")

print("\n" + "=" * 50)
blank_count = sum(1 for i, p in enumerate(reader.pages) 
                  if not bool(re.search(r'[\w\u4e00-\u9fff]', p.extract_text() or "")))
print(f"结论: {total} 页中有 {blank_count} 个疑似空白页")
if blank_count > 0:
    print("如果这些空白页没有被自动删除,")
    print("请把这个输出截图发给开发者!")
print("=" * 50)
input("\n按回车退出...")
