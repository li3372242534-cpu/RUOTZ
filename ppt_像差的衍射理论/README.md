# 像差的衍射理论 — 《光学原理》课程报告

> Diffraction Theory of Aberrations — course presentation for *Principles of Optics*.

## 内容

**`像差的衍射理论_课程报告.pptx`**（21 页，16:9）

- **第一部分 · 基本原理**（公式 + 插图）
  - 几何像差 vs 衍射像差；波像差函数 W 与光瞳函数
  - 衍射积分与点扩散函数 (PSF)；艾里斑；各类像差对 PSF 的影响
  - Zernike 圆多项式；Nijboer–Zernike 衍射理论
  - 斯特列尔比与 Maréchal 判据；MTF；扩展 Nijboer–Zernike (ENZ)
- **第二部分 · 前沿文献调研**（按“讲文献”方式，配原文图）
  - 文献 A (arXiv:2604.26371, 2026)：将 Nijboer–Zernike 理论用于**空间引力波探测**中远场波前误差与 TTL 噪声的解析建模
  - 文献 B (arXiv:2404.15231, 2024)：用 **深度学习 (ResNet-50)** 从离焦 PSF 图像直接预测 Zernike 系数，实现快速像差反演
- **第三部分 · 总结与展望**

## 图片说明

- `figs/gen/`：第一部分插图，由 `gen_figs.py` 用 Python (NumPy/SciPy/Matplotlib) 依据理论计算**自绘**（PSF 由光瞳函数 FFT 得到，曲线由解析公式得到）。
- `figs/paperA/`、`figs/paperB/`：第二部分插图，**截取自对应文献**（arXiv 公开预印本，仅用于课程学习，已在 PPT 中标注来源）。

## 复现

```bash
pip install python-pptx matplotlib scipy pillow numpy
python gen_figs.py     # 重新生成第一部分插图
python build_ppt.py    # 重新生成 PPT
```

## 参考文献

1. M. Born & E. Wolf, *Principles of Optics*, Ch. 9.
2. B. R. A. Nijboer, *Physica* **10** (1943) 679; **13** (1947) 605.
3. A. J. E. M. Janssen, *JOSA A* **19**, 849 (2002); J. Braat et al., *JOSA A* **19**, 858 (2002).
4. Tao, Gao, Xu, Wu, arXiv:2604.26371 (2026). [CC BY 4.0]
5. Kok, Bentley, Parkes, Wright, Somekh, Pound, arXiv:2404.15231 (2024).
