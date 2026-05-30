# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['format_conversion.py'],
    pathex=[],
    binaries=[],
    datas=[('latex_to_omml.py', '.'), ('table_processor.py', '.'), ('sanitize.py', '.'), ('template_config.py', '.'), ('style_shell.py', '.'),
        ('core/__init__.py', 'core'), ('core/latex_to_text.py', 'core'), ('core/numbering_engine.py', 'core'), ('core/ai_normalizer.py', 'core'), ('core/docx_reader.py', 'core'), ('core/ole_washer.py', 'core'), ('core/formula_normalize.py', 'core'), ('core/formula_stats.py', 'core'), ('core/repair.py', 'core'), ('core/escape_features.py', 'core'), ('core/semantics.py', 'core'), ('core/symbols.py', 'core'), ('core/config_manager.py', 'core'), ('core/report/collector.py', 'core/report'), ('core/report/markdown_report.py', 'core/report'), ('core/scene/schema.py', 'core/scene'), ('core/scene/manager.py', 'core/scene'), ('presets/default_format.json', 'presets'), ('presets/thesis_format.json', 'presets'),
        ('splash.py', '.'), ('docx_renderer.py', '.'), ('analyzer/__init__.py', 'analyzer'), ('analyzer/doc_tree.py', 'analyzer'), ('analyzer/change_tracker.py', 'analyzer'),
        ('md_parser/__init__.py', 'markdown'), ('md_parser/ir.py', 'markdown'), ('md_parser/inline_parser.py', 'markdown'), ('md_parser/block_parser.py', 'markdown'), ('md_parser/word_render.py', 'markdown'), ('md_parser/line_spacing.py', 'markdown'),
    ('formula_core/__init__.py', 'formula_core'), ('formula_core/mathtype_ole.py', 'formula_core'),
    ('ai_assist/__init__.py', 'ai_assist'), ('ai_assist/deepseek_client.py', 'ai_assist'), ('ai_assist/image_capture.py', 'ai_assist'), ('ai_assist/settings_dialog.py', 'ai_assist')],
    hiddenimports=['latex_to_omml', 'lxml', 'tkinter', 'requests'],
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
    name='FormatX-CLI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['3.ico'],
)
