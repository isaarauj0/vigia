"""
vigIA — Renderizar slides HTML como PNG e gerar vídeo final
Áudio: NotebookLM (411.4s)
Slides: 10 slides HTML vigIA
"""
import subprocess, os, math, time
from playwright.sync_api import sync_playwright

SCRATCH = "/tmp/claude-0/-home-claude/0f3da570-6e0f-5003-a59a-a6b7c1e79178/scratchpad"
HTML_FILE = f"file://{SCRATCH}/vigia_presentation.html"
AUDIO    = f"{SCRATCH}/audio_notebook.aac"
OUT_DIR  = f"{SCRATCH}/slides_png"
OUTPUT   = f"{SCRATCH}/vigIA_Video_Final.mp4"

AUDIO_DURATION = 411.4  # segundos

# Timings originais → escalar para cobrir 411.4s
TIMINGS_ORIG = [20, 45, 30, 40, 30, 35, 40, 25, 25, 25]  # = 315s
SCALE = AUDIO_DURATION / sum(TIMINGS_ORIG)
TIMINGS_SCALED = [t * SCALE for t in TIMINGS_ORIG]

print(f"Escala: {SCALE:.3f}")
for i, t in enumerate(TIMINGS_SCALED):
    print(f"  Slide {i+1}: {TIMINGS_ORIG[i]}s → {t:.1f}s")
print(f"Total: {sum(TIMINGS_SCALED):.1f}s")

os.makedirs(OUT_DIR, exist_ok=True)

# Renderizar cada slide
with sync_playwright() as p:
    browser = p.chromium.launch(
        executable_path="/opt/pw-browsers/chromium",
        args=["--no-sandbox", "--disable-dev-shm-usage"]
    )
    page = browser.new_page(viewport={"width": 1280, "height": 720})
    page.goto(HTML_FILE, wait_until="networkidle")
    time.sleep(1.5)

    # Para o auto-advance
    page.evaluate("clearInterval(window._timer); window._timer = null;")

    for i in range(10):
        # Vai direto para o slide i
        page.evaluate(f"""
            // Esconde todos os slides
            document.querySelectorAll('.slide').forEach(s => s.classList.remove('active'));
            // Ativa o slide {i}
            const slides = document.querySelectorAll('.slide');
            slides[{i}].classList.add('active');
            // Força animações a aparecerem já
            slides[{i}].querySelectorAll('[class*=" a"]').forEach(el => {{
                el.style.opacity = '1';
                el.style.transform = 'none';
                el.style.animation = 'none';
            }});
        """)
        time.sleep(0.6)

        path = f"{OUT_DIR}/slide_{i+1:02d}.png"
        page.screenshot(path=path, full_page=False)
        print(f"  ✓ slide {i+1:02d}.png capturado")

    browser.close()

print("\nTodos os slides capturados. Montando concat list...")

# Gera arquivo de lista para ffmpeg concat
concat_file = f"{SCRATCH}/concat_list.txt"
with open(concat_file, "w") as f:
    for i, dur in enumerate(TIMINGS_SCALED):
        f.write(f"file '{OUT_DIR}/slide_{i+1:02d}.png'\n")
        f.write(f"duration {dur:.3f}\n")
    # Repete o último slide (necessário para ffmpeg)
    f.write(f"file '{OUT_DIR}/slide_10.png'\n")

print(f"Concat list gerado: {concat_file}")

# Montar vídeo com ffmpeg
print("\nMontando vídeo final com ffmpeg...")
cmd = [
    "ffmpeg", "-y",
    "-f", "concat", "-safe", "0", "-i", concat_file,
    "-i", AUDIO,
    "-vf", "scale=1280:720:force_original_aspect_ratio=decrease,pad=1280:720:(ow-iw)/2:(oh-ih)/2:color=#070A12",
    "-c:v", "libx264", "-preset", "fast", "-crf", "18",
    "-c:a", "aac", "-b:a", "128k",
    "-shortest",
    "-pix_fmt", "yuv420p",
    OUTPUT
]
result = subprocess.run(cmd, capture_output=True, text=True)
if result.returncode == 0:
    size_mb = os.path.getsize(OUTPUT) / 1024 / 1024
    print(f"\n✅ Vídeo gerado: {OUTPUT}")
    print(f"   Tamanho: {size_mb:.1f} MB")
else:
    print("ERRO ffmpeg:")
    print(result.stderr[-2000:])
