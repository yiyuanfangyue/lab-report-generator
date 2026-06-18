# -*- coding: utf-8 -*-
"""
阶段 2：截图分析
================
分析实验截图，建立"截图 → 报告中位置"的映射。

流程：
  1. 遍历截图目录，收集所有图片
  2. 调用 vision.js 识别每张截图的场景内容
  3. 与实验内容结构中的"图片插入点"进行语义匹配
  4. 输出 image_map.json（含置信度）

支持无 vision.js 时的回退方案：用户手动输入描述。
"""
import sys, os, json, subprocess, re
from lib.docx_utils import save_json, find_vision_js, load_json


# -- 中文关键词 → 图片插入点类型 --
TYPE_KEYWORDS = {
    '配置': 'config',
    '注册': 'register',
    '会话': 'session',
    '连接': 'connect',
    '语音': 'voice',
    '视频': 'video',
    '直播': 'live',
    '路测': 'dttest',
    '优化': 'optimize',
    '波束': 'beam',
    '切片': 'slice',
    '自动驾驶': 'auto_drive',
    'V2X': 'v2x',
    '终端': 'terminal',
    'UDM': 'udm',
    '铁塔': 'tower',
    '扇区': 'sector',
    '覆盖': 'coverage',
}


def scan_images(directory):
    """扫描目录，返回图片文件列表"""
    exts = {'.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff', '.webp'}
    images = []
    for fname in sorted(os.listdir(directory)):
        ext = os.path.splitext(fname)[1].lower()
        if ext in exts:
            images.append(os.path.join(directory, fname))
    return images


def analyze_with_vision(image_path, vision_js_path, question=None):
    """调用 vision.js 分析单张截图（10s 超时）"""
    if not vision_js_path or not os.path.exists(vision_js_path):
        return None

    q = question or '请分析这张截图的场景内容，描述图中有什么界面、菜单、按钮、数据配置等信息。如果是实验平台截图，请指出所属的实验环节。'
    try:
        result = subprocess.run(
            ['node', vision_js_path, image_path, q],
            capture_output=True, text=True, timeout=10, encoding='utf-8'
        )
        if result.returncode == 0 and result.stdout:
            return result.stdout.strip()
        else:
            err = result.stderr[:100] if result.stderr else 'unknown error'
            print(f'  [!] vision.js error: {err}')
            return None
    except subprocess.TimeoutExpired:
        print(f'  [!] vision.js 超时（10s）')
        return None
    except FileNotFoundError:
        print(f'  [!] 未找到 node.js')
        return None


def match_description_to_insertions(description, insertions):
    """
    将 vision 识别出的描述文本与插入点列表进行语义匹配。
    返回匹配结果列表（按置信度排序）。
    """
    if not description or not insertions:
        return []

    desc_lower = description.lower()
    scores = []

    for ins in insertions:
        score = 0
        ctx = ins.get('context', '') + ' ' + ins.get('description', '')
        ctx_lower = ctx.lower()

        # 1. 类型关键词匹配
        for kw, kw_type in TYPE_KEYWORDS.items():
            if kw in desc_lower and kw in ctx_lower:
                score += 20

        # 2. 共享的关键词
        # 提取中英文单词/词组
        tokens = set(re.findall(r'[一-鿿\w]+', desc_lower))
        ctx_tokens = set(re.findall(r'[一-鿿\w]+', ctx_lower))
        common = tokens & ctx_tokens
        score += len(common) * 2

        # 3. 数字/编号匹配
        nums_desc = set(re.findall(r'\d+', desc_lower))
        nums_ctx = set(re.findall(r'\d+', ctx_lower))
        common_nums = nums_desc & nums_ctx
        score += len(common_nums) * 5

        # 4. 章节归属匹配
        sec_title = ins.get('section_title', '')
        if sec_title:
            sec_tokens = set(re.findall(r'[一-鿿\w]+', sec_title))
            common_sec = tokens & sec_tokens
            score += len(common_sec) * 3

        scores.append({
            'insertion_id': ins['id'],
            'context': ctx[:80],
            'score': score,
            'section_title': ins.get('section_title', ''),
        })

    # 按分数排序
    scores.sort(key=lambda x: x['score'], reverse=True)
    return scores


def manual_input_prompt(image_path):
    """用户手动输入图片描述（当 vision 不可用或匹配失败时）"""
    fname = os.path.basename(image_path)
    print(f'\n  [IMG] {fname}')
    desc = input('  请输入这张截图的描述（或回车跳过）: ').strip()
    return desc if desc else None


def guess_insertion_position(filename, image_description, insertions):
    """
    综合文件名和描述，猜测应插入的位置。
    """
    # 尝试从文件名中提取提示信息
    fname = os.path.splitext(os.path.basename(filename))[0]

    # 匹配插入点
    text_to_match = fname + ' ' + (image_description or '')
    matches = match_description_to_insertions(text_to_match, insertions)

    if matches:
        best = matches[0]
        if best['score'] >= 10:
            return best['insertion_id'], best['score']

    return None, 0


def analyze_images(image_dir, content_json_path=None, vision_js_path=None):
    """主函数：分析截图"""
    # -- 加载插入点信息 --
    insertions = []
    if content_json_path and os.path.exists(content_json_path):
        content = load_json(content_json_path)
        insertions = content.get('image_insertions', [])
        print(f'  已加载 {len(insertions)} 个图片插入点')

    # -- 扫描截图 --
    images = scan_images(image_dir)
    print(f'  找到 {len(images)} 张截图')

    # -- 检查 vision.js 是否可用（API key 是否配置） --
    vision_available = False
    if vision_js_path and os.path.exists(vision_js_path):
        if os.environ.get('DASHSCOPE_API_KEY'):
            vision_available = True
            print(f'  使用 vision.js: {vision_js_path}')
        else:
            print(f'  已找到 vision.js，但未设置 DASHSCOPE_API_KEY 环境变量')
            print(f'  将使用手动输入模式')
    else:
        print(f'  未找到 vision.js，使用手动输入模式')

    # -- 分析每张截图 --
    image_map = {}
    unmatched = []

    for i, img_path in enumerate(images):
        fname = os.path.basename(img_path)
        print(f'\n  [{i+1}/{len(images)}] 分析: {fname}')

        # 调用 vision（仅当 API key 配置了）
        description = None
        if vision_available:
            description = analyze_with_vision(img_path, vision_js_path)

        if description:
            short_desc = description[:100].replace('\n', ' ')
            print(f'    → {short_desc}...')
        else:
            # 回退到手动输入
            description = manual_input_prompt(img_path) or f'截图 {fname}'

        # 尝试匹配位置
        matched_id, score = guess_insertion_position(fname, description, insertions)

        if matched_id and score >= 10:
            image_map[matched_id] = img_path
            print(f'    [OK] 匹配 → [{matched_id}] (置信度: {score})')
        else:
            print(f'    ? 未能自动匹配')
            unmatched.append({
                'path': img_path,
                'filename': fname,
                'description': description,
                'suggestions': match_description_to_insertions(description, insertions)[:3]
                if insertions else [],
            })

    # -- 处理未匹配的截图 --
    if unmatched and insertions:
        print(f'\n{"="*50}')
        print(f'  以下 {len(unmatched)} 张截图未能自动匹配，请输入对应插入点 ID：')
        print(f'  可用插入点:')
        for ins in insertions:
            used = '[OK]' if ins['id'] in image_map else ' '
            print(f'    [{used}] {ins["id"]}: {ins["context"][:60]}')

        for item in unmatched:
            print(f'\n  📷 {item["filename"]}')
            print(f'     描述: {item["description"][:80]}')
            if item['suggestions']:
                print(f'     建议: {item["suggestions"][0]["insertion_id"]} '
                      f'(分数: {item["suggestions"][0]["score"]})')
            ans = input(f'     请输入插入点 ID（回车跳过）: ').strip()
            if ans in [ins['id'] for ins in insertions]:
                image_map[ans] = item['path']
                print(f'    [OK] 已映射到 {ans}')

    # -- 补齐未匹配的插入点 --
    result_map = {}
    all_ids = [ins['id'] for ins in insertions]
    for ins in insertions:
        iid = ins['id']
        if iid in image_map:
            result_map[iid] = image_map[iid]
        else:
            result_map[iid] = ''

    # -- 检查多余图片 --
    extra = []
    for iid, path in image_map.items():
        if path and iid not in all_ids:
            extra.append({'path': path, 'suggested_id': iid})

    result = {
        'image_map': result_map,
        'matched_count': sum(1 for v in result_map.values() if v),
        'unmatched_count': sum(1 for v in result_map.values() if not v),
        'extra_images': extra,
    }

    return result


def list_images_for_model(image_dir, content_json_path=None, output_map_path='image_map.json'):
    """
    --model-vision 模式：
    当 AI 模型自身具备读图能力（如 GPT-4V、Claude Vision 等）时使用。
    仅列出所有截图和插入点，不调用 vision.js，由模型自行用 Read 工具读图后建立映射。
    输出一个空的 image_map.json 模板供模型填写。
    """
    # 保护：如果已有有效映射，不覆盖
    if os.path.exists(output_map_path):
        try:
            existing = load_json(output_map_path)
            existing_map = existing.get('image_map', existing)
            filled = [v for v in existing_map.values() if v and (isinstance(v, str) or v)]
            if filled:
                print(f'  [!] 警告: {output_map_path} 已有 {len(filled)} 个有效映射，不会被覆盖')
                print(f'  [!] 如需重新映射，请先删除或清空 {output_map_path}')
                # 仍然输出列表供参考，但不覆盖文件
                _list_images_only(image_dir, content_json_path)
                return existing
        except Exception:
            pass

    _list_images_only(image_dir, content_json_path)

    # 生成空模板
    insertions = []
    if content_json_path and os.path.exists(content_json_path):
        content = load_json(content_json_path)
        insertions = content.get('image_insertions', [])

    template_map = {}
    for ins in insertions:
        iid = ins['id']
        cnt = ins.get('expected_count', 1)
        if cnt == 1:
            template_map[iid] = ''
        else:
            template_map[iid] = []

    result = {
        'image_map': template_map,
        'matched_count': 0,
        'unmatched_count': sum(ins.get('expected_count', 1) for ins in insertions),
        'extra_images': [],
    }
    save_json(result, output_map_path)
    print(f'  空映射模板已输出到: {output_map_path}')
    print(f'  请用你的 Read 工具查看截图后，编辑 {output_map_path} 填写实际路径。')
    return result


def _list_images_only(image_dir, content_json_path=None):
    """仅列出截图和插入点，不写文件"""
    images = scan_images(image_dir)
    print(f'  找到 {len(images)} 张截图')
    print()

    print('=' * 60)
    print('  截图列表（供 AI 模型逐一读图识别）：')
    print('=' * 60)
    for i, img_path in enumerate(images):
        fname = os.path.basename(img_path)
        size_kb = os.path.getsize(img_path) // 1024
        print(f'  [{i+1:2d}] {fname}  ({size_kb} KB)')
    print()

    insertions = []
    if content_json_path and os.path.exists(content_json_path):
        content = load_json(content_json_path)
        insertions = content.get('image_insertions', [])
    if insertions:
        print('=' * 60)
        print('  报告中的图片插入点（需将截图映射到这里）：')
        print('=' * 60)
        for ins in insertions:
            used = ins.get('expected_count', 1)
            print(f'  [{ins["id"]}] 需 {used} 张 | {ins["context"][:70]}')

    print()
    print('=' * 60)
    print('  操作说明：')
    print('  AI 模型应用 Read 工具逐张查看截图，判断每张截图的内容，')
    print('  然后手动建立映射关系，输出到 image_map.json。')
    print('  - 单图映射：{"img_01": "路径"}')
    print('  - 多图映射：{"img_03": ["路径1", "路径2", ...]}')
    print('=' * 60)


# ======================================================================
if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('用法: python analyze_images.py <截图目录> [实验内容结构.json] [输出路径]')
        print('       python analyze_images.py --model-vision <截图目录> [实验内容结构.json] [输出路径]')
        print('说明: 分析截图，建立与实验内容的映射关系')
        print('  --model-vision  模型自身读图模式（无需 vision.js，AI 用 Read 工具自行读图）')
        sys.exit(1)

    model_vision_mode = '--model-vision' in sys.argv
    args = [a for a in sys.argv[1:] if a != '--model-vision']

    if model_vision_mode:
        image_dir = args[0] if len(args) > 0 else '.'
        content_json = args[1] if len(args) > 1 else '实验内容结构.json'
        output_path = args[2] if len(args) > 2 else 'image_map.json'
        if not os.path.isdir(image_dir):
            print(f'错误: 目录不存在 — {image_dir}')
            sys.exit(1)
        list_images_for_model(image_dir, content_json if os.path.exists(content_json) else None, output_path)
        sys.exit(0)

    # 原有逻辑
    image_dir = sys.argv[1]
    if not os.path.isdir(image_dir):
        print(f'错误: 目录不存在 — {image_dir}')
        sys.exit(1)

    content_json = sys.argv[2] if len(sys.argv) > 2 else '实验内容结构.json'
    output_path = sys.argv[3] if len(sys.argv) > 3 else 'image_map.json'

    if not os.path.exists(content_json):
        print(f'[!] 未找到实验内容结构文件 {content_json}，将仅收集截图列表')
        content_json = None

    result = analyze_images(image_dir, content_json)

    save_json(result, output_path)

    print(f'\n{"="*50}')
    print(f'[OK] 截图分析完成')
    print(f'  已匹配: {result["matched_count"]}')
    print(f'  未匹配: {result["unmatched_count"]}')
    if result['extra_images']:
        print(f'  多余截图: {len(result["extra_images"])} 张')
    print(f'  输出: {output_path}')
