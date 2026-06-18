# -*- coding: utf-8 -*-
"""
阶段 0：实验内容分析
====================
从实验内容 .docx 中提取结构化内容（章节、正文、图片占位），
实验指导书仅作为补充参考。

交互流程：
  1. 解析文档 → 展示内容树 → 让用户确认/修正
  2. 可选：用实验指导书补充内容
  3. 输出实验内容结构.json

用法:
  python analyze_content.py <实验内容.docx> [输出路径]
  python analyze_content.py --guidebook <实验指导书> <实验内容.docx> [输出路径]
"""
import sys, os, re, json
from lib.docx_utils import save_json


def parse_template(template_path):
    """
    从模板 .docx 提取内容结构（基于样式 ID）。
    style 映射: 2=heading1, 3=heading2, 4=heading3
    """
    from docx import Document
    from docx.oxml.ns import qn

    doc = Document(template_path)
    sections = []
    image_insertions = []
    img_counter = 0

    # 样式名 → 实际类型映射
    style_map = {}
    for s in doc.styles:
        if s.style_id in ('2', '3', '4'):
            e = s.element
            rPr = e.find(qn('w:rPr'))
            sz = None
            if rPr is not None:
                sz_el = rPr.find(qn('w:sz'))
                if sz_el is not None:
                    sz = int(sz_el.get(qn('w:val'))) // 2
            style_map[s.style_id] = {
                'name': s.name,
                'size': sz,
            }

    # 提取封面内容（直到第一个 heading 之前）
    cover_section = {'type': 'cover', 'children': []}
    in_cover = True
    current_h1 = None
    current_h2 = None
    current_h3 = None

    for p in doc.paragraphs:
        text = p.text
        style_id = p.style.style_id if p.style else ''

        # 检测封面结束：遇到 style=2/3/4 即进入正文
        if in_cover:
            if style_id in ('2', '3', '4'):
                in_cover = False
            else:
                # 检测图片
                has_img = bool(p._element.findall(qn('w:drawing')))
                cover_section['children'].append({
                    'type': 'image' if has_img else 'body',
                    'text': text,
                    'style_id': style_id,
                })
                continue

        # ===== 正文区域 =====
        entry = {'text': text, 'style_id': style_id}

        if style_id == '2':  # heading 1（模板中只有"实验内容"）
            current_h1 = {
                'type': 'heading1',
                'title': text,
                'children': [],
            }
            sections.append(current_h1)
            current_h2 = None
            current_h3 = None

        elif style_id == '3':  # heading 2（实验目的、实验原理等）
            current_h2 = {
                'type': 'heading2',
                'title': text,
                'children': [],
                'style': style_map.get('3', {}),
            }
            if current_h1:
                current_h1['children'].append(current_h2)
            else:
                sections.append(current_h2)
            current_h3 = None

        elif style_id == '4':  # heading 3（作业题、子任务标题）
            current_h3 = {
                'type': 'heading3',
                'title': text,
                'children': [],
                'style': style_map.get('4', {}),
            }
            if current_h2:
                current_h2['children'].append(current_h3)
            elif current_h1:
                current_h1['children'].append(current_h3)
            else:
                sections.append(current_h3)

        else:  # body / 列表 / 图片占位
            parent = current_h3 or current_h2 or current_h1
            if not parent:
                parent = {'children': sections}
                if not hasattr(parent, 'children'):
                    # Hmm, sections is a list
                    pass

            # 判断图片占位
            is_img_placeholder = False
            if any(kw in text for kw in ['截图', '粘贴']):
                if re.search(r'[①②③]', text):
                    is_img_placeholder = True

            if is_img_placeholder:
                img_counter += 1
                img_id = f'img_{img_counter:02d}'
                m = re.search(r'共计\s*(\d+)\s*张', text)
                expected = int(m.group(1)) if m else 1
                image_insertions.append({
                    'id': img_id,
                    'context': text,
                    'section': current_h2['title'] if current_h2 else (current_h1['title'] if current_h1 else ''),
                    'expected_count': expected,
                })
                child = {'type': 'image_placeholder', 'id': img_id, 'text': text, 'expected_count': expected}
            else:
                child = {'type': 'body', 'text': text}

            if isinstance(parent, dict) and 'children' in parent:
                parent['children'].append(child)
            else:
                sections.append(child)

    # 检测连续空行区域作为图片预留区
    # 简化处理：不在模板中检测空行，只依赖显式的"①粘贴"标记

    result = {
        'experiment_name': '',
        'source_file': os.path.basename(template_path),
        'cover': cover_section,
        'sections': sections,
        'image_insertions': image_insertions,
        'style_map': style_map,
    }

    return result


def print_content_tree(result):
    """打印内容树供用户确认"""
    print('\n' + '=' * 60)
    print('  提取的内容结构：')
    print('=' * 60)

    # 封面
    cover = result.get('cover', {})
    cover_kids = cover.get('children', [])
    print(f'\n  [封面] 共 {len(cover_kids)} 段')
    for c in cover_kids[:5]:
        t = c.get('text', '')[:50]
        print(f'    {t}')
    if len(cover_kids) > 5:
        print(f'    ...（共 {len(cover_kids)} 段）')

    # 正文
    print(f'\n  正文章节：')
    for sec in result.get('sections', []):
        _print_node(sec, 0)


def _print_node(node, indent):
    prefix = '  ' * (indent + 1)
    t = node.get('type', '')
    title = node.get('title', node.get('text', ''))[:60]
    if t == 'heading1':
        print(f'{prefix}[H1] {title}')
    elif t == 'heading2':
        print(f'{prefix}[H2] {title}')
    elif t == 'heading3':
        print(f'{prefix}[H3] {title}')
    elif t == 'image_placeholder':
        print(f'{prefix}[IMG] {title}')
    elif t == 'body':
        if title:
            print(f'{prefix}{title}')
    for child in node.get('children', []):
        _print_node(child, indent + 1)


def print_style_info(result):
    """打印样式信息供用户确认"""
    print('\n  检测到的标题样式：')
    for sid, info in result.get('style_map', {}).items():
        print(f'    style={sid}: {info["name"]}, 字号={info["size"]}pt')
    print()


def interactive_confirm(result):
    """交互式确认：让用户修改提取的内容"""
    print_content_tree(result)
    print_style_info(result)

    ans = input('  以上内容结构是否正确？(y=继续, n=放弃, e=编辑样式字号): ').strip().lower()
    if ans == 'n':
        print('  已放弃，请检查模板文件后重试')
        return False
    elif ans == 'e':
        for sid, info in result.get('style_map', {}).items():
            new_sz = input(f'    style={sid} ({info["name"]}) 当前字号={info["size"]}pt, 修改为: ').strip()
            if new_sz:
                try:
                    info['size'] = int(new_sz)
                except ValueError:
                    pass
        print('  样式已更新')
        return True
    return True


# ======================================================================
if __name__ == '__main__':
    # 检查 --yes 标志
    auto_confirm = '--yes' in sys.argv
    args = [a for a in sys.argv[1:] if a != '--yes']

    if not args:
        print('用法: python analyze_content.py [--yes] <实验内容.docx> [输出路径]')
        print('       python analyze_content.py --guidebook <实验指导书> <实验内容.docx> [输出路径]')
        sys.exit(1)

    # 解析参数
    guidebook_path = None

    if '--guidebook' in args:
        idx = args.index('--guidebook')
        guidebook_path = args[idx + 1]
        args = args[:idx] + args[idx + 2:]

    template_path = args[0]
    if not os.path.exists(template_path):
        print(f'错误: 模板文件不存在 — {template_path}')
        sys.exit(1)

    output_path = args[1] if len(args) > 1 else '实验内容结构.json'

    if not os.path.exists(template_path):
        print(f'错误: 模板文件不存在 — {template_path}')
        sys.exit(1)

    if guidebook_path and not os.path.exists(guidebook_path):
        print(f'警告: 指导书文件不存在 — {guidebook_path}')
        guidebook_path = None

    # 解析模板内容
    result = parse_template(template_path)

    if guidebook_path:
        # 读取指导书作为补充参考（仅提取章节结构用于对比）
        print(f'  实验指导书（参考）: {guidebook_path}')
        # 暂时仅标记，内容还是以模板为准

    # 交互确认
    ok = True
    if not auto_confirm:
        ok = interactive_confirm(result)
    if not ok:
        sys.exit(1)

    save_json(result, output_path)
    print(f'[OK] 实验内容分析完成')
    print(f'  章节数: {len(result["sections"])}')
    print(f'  图片插入点: {len(result["image_insertions"])}')
    for img in result['image_insertions']:
        print(f'    [{img["id"]}] {img["context"][:60]}')
    print(f'  输出: {output_path}')
