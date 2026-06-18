# -*- coding: utf-8 -*-
"""
实验室报告生成器 — 通用 4 阶段流水线

用法:
  python analyze_content.py <实验指导书路径> [输出路径]
  python analyze_template.py <模板.docx> [输出前缀]
  python analyze_images.py <截图目录> [实验内容结构.json] [输出路径]
  python generate_report.py --content <实验内容结构.json> --latex-format <LaTeX格式模板.tex>
                            --images <image_map.json> [--template 模板.docx]
                            [--format pdf|docx|both]

依赖:
  pip install python-docx pillow
"""
