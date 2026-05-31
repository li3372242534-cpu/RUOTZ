"""
Generate illustrative figures for the principles section of the PPT
'Diffraction Theory of Aberrations'.
All physics-accurate, computed from first principles (FFT of pupil function,
Zernike polynomials, Bessel functions). Figure labels in English; the Chinese
narration lives in the PPT body text.
"""
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scipy.special import jn
import os

OUT = "/projects/sandbox/ppt_aberration/figs/gen"
os.makedirs(OUT, exist_ok=True)

plt.rcParams.update({
    "figure.dpi": 150,
    "savefig.dpi": 150,
    "font.size": 12,
    "axes.titlesize": 13,
    "axes.labelsize": 12,
    "mathtext.fontset": "cm",
    "font.family": "DejaVu Sans",
    "savefig.bbox": "tight",
    "savefig.facecolor": "white",
})

INK = "#1b2a4a"
ACC = "#c0392b"
ACC2 = "#1f77b4"
GREEN = "#2e8b57"


def zernike(name, rho, theta):
    if name == "defocus":
        return np.sqrt(3) * (2 * rho**2 - 1)
    if name == "astig":
        return np.sqrt(6) * rho**2 * np.cos(2 * theta)
    if name == "coma":
        return np.sqrt(8) * (3 * rho**3 - 2 * rho) * np.cos(theta)
    if name == "trefoil":
        return np.sqrt(8) * rho**3 * np.cos(3 * theta)
    if name == "spherical":
        return np.sqrt(5) * (6 * rho**4 - 6 * rho**2 + 1)
    if name == "tiltx":
        return 2 * rho * np.cos(theta)
    raise ValueError(name)


def disk_grid(N=256):
    x = np.linspace(-1, 1, N)
    X, Y = np.meshgrid(x, x)
    R = np.hypot(X, Y)
    T = np.arctan2(Y, X)
    mask = R <= 1.0
    return X, Y, R, T, mask


def fig_wavefront_schematic():
    fig, ax = plt.subplots(figsize=(7.6, 4.6))
    ax.plot([0, 0], [-2.3, 2.3], color="#888", lw=1.2)
    ax.text(0.05, 2.32, "Exit pupil plane", color="#555", fontsize=11)
    xc = 6.0
    ax.plot(xc, 0, "o", color=INK, ms=7)
    ax.text(xc + 0.12, 0.05, "Gaussian\nimage point $P_0$", color=INK, fontsize=10, va="center")
    Rsph = xc
    th = np.linspace(-np.deg2rad(23), np.deg2rad(23), 200)
    xs = xc - Rsph * np.cos(th)
    ys = Rsph * np.sin(th)
    ax.plot(xs, ys, color=ACC2, lw=2.4, label="Reference sphere (ideal)")
    deform = 0.55 * np.sin(2.7 * th / th.max() * np.pi) * (np.abs(th) / th.max())
    xa = xs - deform * np.cos(th)
    ya = ys - deform * np.sin(th)
    ax.plot(xa, ya, color=ACC, lw=2.6, label="Actual (aberrated) wavefront")
    i = 150
    ax.annotate("", xy=(xa[i], ya[i]), xytext=(xs[i], ys[i]),
                arrowprops=dict(arrowstyle="<->", color=GREEN, lw=2))
    ax.text(xa[i] - 1.15, ya[i] + 0.18, r"$W(x,y)$", color=GREEN, fontsize=15)
    for yy in np.linspace(-1.8, 1.8, 7):
        ax.plot([0, xc], [yy, 0], color="#bbb", lw=0.7, zorder=0)
    ax.set_xlim(-1.0, 7.6)
    ax.set_ylim(-2.6, 2.8)
    ax.axis("off")
    ax.legend(loc="upper left", fontsize=10, frameon=False)
    ax.set_title(r"Wave aberration $W$: deviation of the real wavefront from the reference sphere",
                 fontsize=12, color=INK)
    fig.savefig(f"{OUT}/01_wavefront_schematic.png")
    plt.close(fig)


def fig_zernike_maps():
    X, Y, R, T, mask = disk_grid(220)
    items = [
        ("tiltx", r"$Z_1^1$  Tilt"),
        ("defocus", r"$Z_2^0$  Defocus"),
        ("astig", r"$Z_2^2$  Astigmatism"),
        ("coma", r"$Z_3^1$  Coma"),
        ("trefoil", r"$Z_3^3$  Trefoil"),
        ("spherical", r"$Z_4^0$  Spherical"),
    ]
    fig, axes = plt.subplots(2, 3, figsize=(8.4, 5.7))
    for ax, (name, title) in zip(axes.ravel(), items):
        Z = zernike(name, R, T)
        Zm = np.where(mask, Z, np.nan)
        ax.imshow(Zm, extent=[-1, 1, -1, 1], cmap="RdBu_r", origin="lower")
        ax.set_title(title, fontsize=12, color=INK)
        ax.set_xticks([]); ax.set_yticks([])
        ax.add_patch(plt.Circle((0, 0), 1.0, color="k", fill=False, lw=1.0))
    fig.suptitle("Zernike circle polynomials: an orthogonal basis of 'balanced' aberrations",
                 fontsize=12.5, color=INK, y=1.02)
    fig.tight_layout()
    fig.savefig(f"{OUT}/02_zernike_maps.png")
    plt.close(fig)


def compute_psf(W_waves, N=256, pad=6):
    x = np.linspace(-1, 1, N)
    X, Y = np.meshgrid(x, x)
    R = np.hypot(X, Y)
    mask = R <= 1.0
    P = np.zeros((N, N), dtype=complex)
    Wf = np.nan_to_num(W_waves)
    P[mask] = np.exp(1j * 2 * np.pi * Wf[mask])
    M = N * pad
    Pp = np.zeros((M, M), dtype=complex)
    s = (M - N) // 2
    Pp[s:s + N, s:s + N] = P
    amp = np.fft.fftshift(np.fft.fft2(np.fft.ifftshift(Pp)))
    return np.abs(amp) ** 2


def crop_center(a, c):
    s = a.shape[0] // 2
    return a[s - c:s + c, s - c:s + c]


def fig_airy():
    X, Y, R, T, mask = disk_grid(256)
    W0 = np.where(mask, 0.0, np.nan)
    psf = compute_psf(W0, N=256, pad=6)
    psf /= psf.max()
    sub = crop_center(psf, 60)
    v = np.linspace(1e-6, 12, 600)
    airy = (2 * jn(1, v) / v) ** 2
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(8.6, 3.9),
                                 gridspec_kw={"width_ratios": [1, 1.25]})
    a1.imshow(sub ** 0.35, cmap="inferno", origin="lower")
    a1.set_title("Airy disk (ideal PSF)", color=INK)
    a1.set_xticks([]); a1.set_yticks([])
    a2.plot(v, airy, color=ACC, lw=2.2)
    a2.fill_between(v, airy, 0, color=ACC, alpha=0.12)
    a2.set_xlabel(r"$v=\frac{2\pi}{\lambda}\frac{a}{R}\,r$  (normalized radius)")
    a2.set_ylabel(r"Intensity  $I/I_0$")
    a2.set_title(r"$I(v)=\left[\,2J_1(v)/v\,\right]^2$", color=INK)
    a2.axvline(3.832, color=ACC2, ls="--", lw=1.2)
    a2.text(3.95, 0.55, "1st dark ring\n$v=3.832$", color=ACC2, fontsize=9)
    a2.set_xlim(0, 12); a2.set_ylim(-0.02, 1.02)
    fig.tight_layout()
    fig.savefig(f"{OUT}/03_airy.png")
    plt.close(fig)


def fig_aberrated_psf():
    X, Y, R, T, mask = disk_grid(256)
    rms = 0.25
    cases = [
        ("Aberration-free", np.zeros_like(R)),
        ("Defocus", rms * zernike("defocus", R, T)),
        ("Astigmatism", rms * zernike("astig", R, T)),
        ("Coma", rms * zernike("coma", R, T)),
        ("Spherical", rms * zernike("spherical", R, T)),
    ]
    fig, axes = plt.subplots(1, 5, figsize=(11.2, 2.7))
    for ax, (title, W) in zip(axes, cases):
        Wm = np.where(mask, W, np.nan)
        psf = compute_psf(Wm, N=256, pad=6)
        psf /= psf.max()
        ax.imshow(crop_center(psf, 55) ** 0.35, cmap="inferno", origin="lower")
        ax.set_title(title, fontsize=11, color=INK)
        ax.set_xticks([]); ax.set_yticks([])
    fig.suptitle(r"Effect of each aberration on the diffraction PSF  (RMS = 0.25$\lambda$)",
                 fontsize=12.5, color=INK, y=1.06)
    fig.tight_layout()
    fig.savefig(f"{OUT}/04_aberrated_psf.png")
    plt.close(fig)


def fig_strehl():
    sigma = np.linspace(0, 0.18, 400)
    phi = 2 * np.pi * sigma
    S_exact = np.exp(-phi ** 2)
    S_quad = 1 - phi ** 2
    fig, ax = plt.subplots(figsize=(7.4, 4.5))
    ax.plot(sigma, S_exact, color=ACC, lw=2.6, label=r"$S\approx e^{-(2\pi\sigma_W)^2}$  (Mar\'echal)")
    ax.plot(sigma, S_quad, color=ACC2, lw=1.8, ls="--", label=r"$S\approx 1-(2\pi\sigma_W)^2$")
    ax.axhline(0.8, color=GREEN, lw=1.4, ls=":")
    sig08 = 1.0 / 14
    ax.axvline(sig08, color=GREEN, lw=1.4, ls=":")
    ax.plot(sig08, 0.8, "o", color=GREEN, ms=8)
    ax.annotate(r"Diffraction limit:  $S=0.8 \Leftrightarrow \sigma_W=\lambda/14$",
                xy=(sig08, 0.8), xytext=(0.072, 0.5),
                arrowprops=dict(arrowstyle="->", color=GREEN, lw=1.6),
                color=GREEN, fontsize=11)
    ax.set_xlabel(r"RMS wavefront error  $\sigma_W$  (in units of $\lambda$)")
    ax.set_ylabel(r"Strehl ratio  $S$")
    ax.set_title("Strehl ratio vs. RMS wavefront error", color=INK)
    ax.set_xlim(0, 0.18); ax.set_ylim(0.3, 1.01)
    ax.legend(loc="lower left", fontsize=10, frameon=False)
    ax.grid(alpha=0.25)
    fig.tight_layout()
    fig.savefig(f"{OUT}/05_strehl.png")
    plt.close(fig)


def mtf_from_pupil(W_waves, N=256):
    x = np.linspace(-1, 1, N)
    X, Y = np.meshgrid(x, x)
    R = np.hypot(X, Y)
    mask = R <= 1.0
    P = np.zeros((N, N), dtype=complex)
    P[mask] = np.exp(1j * 2 * np.pi * np.nan_to_num(W_waves)[mask])
    F = np.fft.fft2(P)
    otf = np.abs(np.fft.fftshift(np.fft.ifft2(np.abs(F) ** 2)))
    otf /= otf.max()
    cy = N // 2
    line = otf[cy, cy:]
    freq = np.arange(line.size) / (N / 2.0)
    return freq, line


def fig_mtf():
    X, Y, R, T, mask = disk_grid(256)
    freqc, mc = mtf_from_pupil(np.where(mask, 0.22 * zernike("coma", R, T), 0.0))
    freqs, ms = mtf_from_pupil(np.where(mask, 0.22 * zernike("spherical", R, T), 0.0))
    nu = np.linspace(0, 1, 300)
    mtf_ideal = (2 / np.pi) * (np.arccos(nu) - nu * np.sqrt(1 - nu ** 2))
    fig, ax = plt.subplots(figsize=(7.4, 4.5))
    ax.plot(nu, mtf_ideal, color="k", lw=2.6, label="Aberration-free (diffraction limit)")
    ax.plot(freqc[freqc <= 1], mc[freqc <= 1], color=ACC, lw=2.0, label=r"With coma $0.22\lambda$")
    ax.plot(freqs[freqs <= 1], ms[freqs <= 1], color=ACC2, lw=2.0, label=r"With spherical $0.22\lambda$")
    ax.set_xlabel(r"Normalized spatial frequency  $\nu/\nu_c$")
    ax.set_ylabel("Modulation Transfer Function (MTF)")
    ax.set_title("Aberrations reduce mid/high-frequency contrast", color=INK)
    ax.set_xlim(0, 1); ax.set_ylim(0, 1.02)
    ax.legend(fontsize=10, frameon=False)
    ax.grid(alpha=0.25)
    fig.tight_layout()
    fig.savefig(f"{OUT}/06_mtf.png")
    plt.close(fig)


def fig_nijboer_zernike():
    v = np.linspace(1e-6, 16, 700)
    fig, ax = plt.subplots(figsize=(7.6, 4.5))
    for n, col, lab in [(0, "k", r"$n{=}0$ (Airy): $2J_1(v)/v$"),
                        (2, ACC2, r"$n{=}2$: $2J_3(v)/v$"),
                        (4, ACC, r"$n{=}4$: $2J_5(v)/v$")]:
        ax.plot(v, 2 * jn(n + 1, v) / v, color=col, lw=2.2, label=lab)
    ax.axhline(0, color="#ccc", lw=0.8)
    ax.set_xlabel(r"Normalized radial coordinate  $v$")
    ax.set_ylabel(r"$2\,J_{n+1}(v)/v$")
    ax.set_title("Nijboer-Zernike: each Zernike term's diffraction integral is an analytic Bessel function",
                 fontsize=10.8, color=INK)
    ax.text(0.98, 0.97,
            r"$\int_0^1 R_n^m(\rho)\,J_m(v\rho)\,\rho\,d\rho=(-1)^{\frac{n-m}{2}}\dfrac{J_{n+1}(v)}{v}$",
            transform=ax.transAxes, ha="right", va="top", fontsize=12,
            bbox=dict(boxstyle="round", fc="#f3f4f7", ec="#ccc"))
    ax.legend(fontsize=10, frameon=False, loc="lower right")
    ax.set_xlim(0, 16)
    ax.grid(alpha=0.22)
    fig.tight_layout()
    fig.savefig(f"{OUT}/07_nijboer_zernike.png")
    plt.close(fig)


def fig_through_focus():
    X, Y, R, T, mask = disk_grid(256)
    defs = [-1.0, -0.5, 0.0, 0.5, 1.0]
    sph = 0.15 * zernike("spherical", R, T)
    fig, axes = plt.subplots(1, 5, figsize=(11.2, 2.7))
    for ax, d in zip(axes, defs):
        W = d * zernike("defocus", R, T) + sph
        Wm = np.where(mask, W, np.nan)
        psf = compute_psf(Wm, N=256, pad=6)
        psf /= psf.max()
        ax.imshow(crop_center(psf, 55) ** 0.35, cmap="viridis", origin="lower")
        ax.set_title(f"defocus = {d:+.1f}$\\lambda$", fontsize=10.5, color=INK)
        ax.set_xticks([]); ax.set_yticks([])
    fig.suptitle("Through-focus PSF evolution (with spherical aberration) - the basis of ENZ retrieval",
                 fontsize=11.5, color=INK, y=1.06)
    fig.tight_layout()
    fig.savefig(f"{OUT}/08_through_focus.png")
    plt.close(fig)


if __name__ == "__main__":
    fig_wavefront_schematic()
    fig_zernike_maps()
    fig_airy()
    fig_aberrated_psf()
    fig_strehl()
    fig_mtf()
    fig_nijboer_zernike()
    fig_through_focus()
    print("All figures generated in", OUT)
    print(sorted(os.listdir(OUT)))
