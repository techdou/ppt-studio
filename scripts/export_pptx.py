#!/usr/bin/env python3
"""PPTD → PPTX 本地导出入口（python-pptx 本地引擎，无网络/登录依赖）。

由 export_pptx_local.py 直接渲染，支持 fade 转场。

用法:
  python export_pptx.py <deck.pptd> [--output deck.pptx] [--transition fade|none] [--force]
"""
import export_pptx_local

if __name__ == "__main__":
    import sys
    sys.argv[0] = "export_pptx.py"
    export_pptx_local.main()
