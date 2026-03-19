# bot.py  |  branch: styled
import os, asyncio, tempfile
from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, FSInputFile
from aiogram.filters import CommandStart
from openai import AsyncOpenAI
from pydub import AudioSegment
from dotenv import load_dotenv
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

load_dotenv()

bot = Bot(token=os.getenv("TELEGRAM_TOKEN"))
dp = Dispatcher()
client = AsyncOpenAI(
    api_key=os.getenv("OPENAI_API_KEY"),
    base_url="https://api.proxyapi.ru/openai/v1"
)

# ── Палитра шаблона ────────────────────────────────────────────────
BG_DARK   = RGBColor(0x13, 0x11, 0x1C)
BG_LIGHT  = RGBColor(0xEC, 0xF1, 0xF3)
CLR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
CLR_BLACK = RGBColor(0x00, 0x00, 0x00)
CLR_GRAY  = RGBColor(0x7C, 0x7C, 0x7C)
CLR_ORANGE = RGBColor(0xFF, 0x5C, 0x00)
CLR_PURPLE = RGBColor(0x7B, 0x4C, 0xFF)
CLR_TEAL   = RGBColor(0x02, 0xB2, 0xA9)
CLR_BLUE   = RGBColor(0x45, 0x79, 0xFF)
ACCENT_CYCLE = [CLR_PURPLE, CLR_ORANGE, CLR_TEAL, CLR_BLUE]
FONT = "Helvetica Neue"

SLIDE_W = Inches(10.0)
SLIDE_H = Inches(5.625)
LABEL   = "Москва 2026"

# ── Системный промт (без изменений) ───────────────────────────────
SYSTEM_PROMPT = """
Ты — ассистент для создания презентаций.
Тебе дают расшифровку голосового сообщения.
Твоя задача — структурировать его в Markdown-презентацию.

Правила:
- Используй ТОЛЬКО то, что сказал пользователь. Не додумывай ничего лишнего.
- Каждый слайд: # Заголовок слайда + 3-5 коротких буллетов
- Максимум 10 слайдов
- Первый слайд — титульный с темой и подзаголовком
- Последний слайд — ключевые выводы

Формат вывода — чистый Markdown, ничего лишнего.
"""

transcripts = {}


# ── Утилиты стилизации ─────────────────────────────────────────────

def set_bg(slide, rgb: RGBColor):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb


def add_text(slide, text, left, top, width, height,
             size=18, bold=False, color=CLR_WHITE, align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color


def add_rect(slide, left, top, width, height, color: RGBColor):
    s = slide.shapes.add_shape(1, left, top, width, height)
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()


def add_chrome(slide, slide_num: int):
    """Шапка с датой + номер слайда — на всех слайдах."""
    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(0.38), BG_DARK)
    add_text(slide, LABEL, Inches(0.14), Inches(0.05), Inches(4), Inches(0.32),
             size=10, color=CLR_WHITE)
    add_text(slide, str(slide_num),
             Inches(9.3), Inches(0.03), Inches(0.5), Inches(0.34),
             size=13, bold=True, color=CLR_ORANGE, align=PP_ALIGN.RIGHT)


# ── Два типа слайдов ───────────────────────────────────────────────

def make_title_slide(prs, title: str, slide_num: int):
    """Тёмный фон, крупный заголовок, оранжевый акцент слева."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, BG_DARK)
    add_chrome(slide, slide_num)
    # Вертикальная оранжевая полоска
    add_rect(slide, Inches(0), Inches(0.38), Inches(0.08), Inches(5.245), CLR_ORANGE)
    # Заголовок
    add_text(slide, title,
             Inches(0.22), Inches(1.2), Inches(9.3), Inches(2.0),
             size=44, bold=True, color=CLR_WHITE)


def make_content_slide(prs, title: str, bullets: list, slide_num: int):
    """Светлый фон, заголовок, оранжевая линия, цветные маркеры у буллетов."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, BG_LIGHT)
    add_chrome(slide, slide_num)
    # Заголовок
    add_text(slide, title,
             Inches(0.22), Inches(0.5), Inches(9.3), Inches(0.85),
             size=30, bold=True, color=CLR_BLACK)
    # Линия под заголовком
    add_rect(slide, Inches(0.22), Inches(1.38), Inches(9.1), Inches(0.04), CLR_ORANGE)
    # Буллеты
    gap = Inches(0.78)
    for i, bullet in enumerate(bullets[:5]):
        top = Inches(1.55) + i * gap
        accent = ACCENT_CYCLE[i % 4]
        # Цветной маркер
        add_rect(slide, Inches(0.22), top + Inches(0.1), Inches(0.07), Inches(0.26), accent)
        # Текст
        add_text(slide, bullet,
                 Inches(0.38), top, Inches(9.1), Inches(0.68),
                 size=15, bold=False, color=CLR_BLACK)


# ── Основная функция — логика из оригинала, стиль новый ───────────

def build_pptx(md_content: str, output_path: str):
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    current_title   = None
    current_bullets = []
    slide_num = [1]   # список чтобы менять из flush_slide (closure)

    def flush_slide():
        if current_title is None:
            return
        if not current_bullets:
            make_title_slide(prs, current_title, slide_num[0])
        else:
            make_content_slide(prs, current_title, current_bullets, slide_num[0])
        slide_num[0] += 1

    for line in md_content.splitlines():
        line = line.strip()
        if line.startswith("# "):
            flush_slide()
            current_title   = line[2:].strip()
            current_bullets = []
        elif line.startswith("## "):
            flush_slide()
            current_title   = line[3:].strip()
            current_bullets = []
        elif line.startswith(("- ", "* ", "• ")):
            current_bullets.append(line[2:].strip())

    flush_slide()
    prs.save(output_path)


# ── Telegram-хэндлеры (без изменений) ─────────────────────────────

@dp.message(CommandStart())
async def start(message: Message):
    await message.answer(
        "Привет! 🎤\n\n"
        "Наговори голосовое — я соберу из него презентацию.\n"
        "Можешь отправить несколько войсов подряд, затем напиши /build"
    )


@dp.message(F.voice)
async def handle_voice(message: Message):
    user_id = message.from_user.id
    await message.answer("⏳ Транскрибирую...")

    voice_file = await bot.get_file(message.voice.file_id)
    with tempfile.TemporaryDirectory() as tmp:
        oga_path = f"{tmp}/voice.oga"
        mp3_path = f"{tmp}/voice.mp3"
        await bot.download_file(voice_file.file_path, oga_path)
        AudioSegment.from_ogg(oga_path).export(mp3_path, format="mp3")

        with open(mp3_path, "rb") as f:
            result = await client.audio.transcriptions.create(
                model="whisper-1", file=f, language="ru"
            )

    text = result.text
    transcripts[user_id] = transcripts.get(user_id, "") + "\n\n" + text
    await message.answer(
        f"✅ Записал:\n\n_{text}_\n\nОтправь ещё войс или /build",
        parse_mode="Markdown"
    )


@dp.message(F.text == "/build")
async def build_presentation(message: Message):
    user_id = message.from_user.id
    if user_id not in transcripts or not transcripts[user_id].strip():
        await message.answer("Сначала отправь голосовое сообщение 🎤")
        return

    await message.answer("🔨 Собираю презентацию...")

    response = await client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": transcripts[user_id]}
        ]
    )

    md_content = response.choices[0].message.content

    with tempfile.TemporaryDirectory() as tmp:
        pptx_path = f"{tmp}/presentation.pptx"
        build_pptx(md_content, pptx_path)
        await message.answer_document(
            FSInputFile(pptx_path, filename="presentation.pptx"),
            caption="🎉 Готово! Презентация на основе твоих слов."
        )

    del transcripts[user_id]


async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
