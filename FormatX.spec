# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('3.ico', '.'), ('Montserrat-Bold.ttf', '.'), ('latex_to_omml.py', '.'), ('format_conversion.py', '.'), ('table_processor.py', '.'), ('sanitize.py', '.'), ('template_config.py', '.'), ('style_shell.py', '.'),
    ('core/__init__.py', 'core'), ('core/latex_to_text.py', 'core'), ('core/numbering_engine.py', 'core'), ('core/ai_normalizer.py', 'core'), ('core/docx_reader.py', 'core'), ('core/ole_washer.py', 'core'), ('core/formula_stats.py', 'core'), ('core/report/collector.py', 'core/report'), ('core/report/markdown_report.py', 'core/report'), ('core/scene/__init__.py', 'core/scene'), ('core/scene/schema.py', 'core/scene'), ('core/scene/manager.py', 'core/scene'), ('presets/default_format.json', 'presets'), ('presets/thesis_format.json', 'presets'), ('presets/ieee_conference.json', 'presets'), ('core/repair.py', 'core'), ('core/escape_features.py', 'core'), ('core/semantics.py', 'core'), ('core/symbols.py', 'core'), ('core/title_dictionary.py', 'core'), ('core/prompts.py', 'core'), ('core/formula_normalize.py', 'core'), ('core/config_manager.py', 'core'), ('core/constants.py', 'core'), ('core/heading_numbering.py', 'core'), ('core/utils/__init__.py', 'core/utils'), ('core/utils/chinese.py', 'core/utils'), ('core/indent.py', 'core'), ('core/line_spacing.py', 'core'), ('core/ooxml_paragraph.py', 'core'), ('core/toc_entry.py', 'core'), ('core/runtime_probe.py', 'core'), ('core/advanced_formatter.py', 'core'), ('core/safe_payload_rewriter.py', 'core'), ('core/prompt_compiler.py', 'core'), ('core/state/__init__.py', 'core/state'), ('core/state/constants.py', 'core/state'), ('core/state/mode_manager.py', 'core/state'), ('core/clone_engine/__init__.py', 'core/clone_engine'), ('core/clone_engine/extractor.py', 'core/clone_engine'), ('core/clone_engine/translator.py', 'core/clone_engine'), ('core/clone_engine/api.py', 'core/clone_engine'),
    ('pipeline_adapter.py', '.'), ('splash.py', '.'), ('docx_renderer.py', '.'), ('analyzer/__init__.py', 'analyzer'), ('analyzer/doc_tree.py', 'analyzer'), ('analyzer/change_tracker.py', 'analyzer'),
    ('src/scene/presets/default_format.json', 'src/scene/presets'),
    ('md_parser/__init__.py', 'markdown'), ('md_parser/ir.py', 'markdown'), ('md_parser/inline_parser.py', 'markdown'), ('md_parser/block_parser.py', 'markdown'), ('md_parser/word_render.py', 'markdown'), ('md_parser/line_spacing.py', 'markdown'),
    ('formula_core/__init__.py', 'formula_core'), ('formula_core/mathtype_ole.py', 'formula_core'),
    ('ai_assist/__init__.py', 'ai_assist'), ('ai_assist/deepseek_client.py', 'ai_assist'), ('ai_assist/image_capture.py', 'ai_assist'), ('ai_assist/settings_dialog.py', 'ai_assist'), ('ai_assist/ai_chat_panel.py', 'ai_assist'), ('ai_assist/provider_presets.py', 'ai_assist'), ('ai_assist/multi_llm_client.py', 'ai_assist'), ('ai_assist/stream_buffer.py', 'ai_assist'), ('ai_assist/stream_worker.py', 'ai_assist'), ('ai_assist/block_dispatcher.py', 'ai_assist'), ('ai_assist/typewriter_engine.py', 'ai_assist')]
binaries = []
hiddenimports = ['latex_to_omml', 'lxml', 'customtkinter', 'tkinterdnd2', 'PIL', 'requests']
tmp_ret = collect_all('customtkinter')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('tkinterdnd2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['gui_main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FormatX',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['3.ico'],
)
