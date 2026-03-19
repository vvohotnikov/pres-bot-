# bot.py
import os, asyncio, tempfile
from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, FSInputFile
from aiogram.filters import CommandStart
from openai import AsyncOpenAI
from pydub import AudioSegment
from dotenv import load_dotenv
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

load_dotenv()

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template.pptx")

bot = Bot(token=os.getenv("TELEGRAM_TOKEN"))
dp = Dispatcher()
client = AsyncOpenAI(
    api_key=os.getenv("OPENAI_API_KEY"),
    base_url="https://api.proxyapi.ru/openai/v1"
)

SYSTEM_PROMPT = """
Ты — ассистент для создания презентаций.
Тебе дают расшифровку голосового сообщения.
Твоя задача — структурировать его в Markdown-презентацию.

Правила:
- Используй ТОЛЬКО то, что сказал пользователь. Не додумывай ничего лишнего.
- Каждый слайд: # Заголовок слайда + 3-5 коротких буллетов
- Максимум 10 слайдов
- Первый слайд — титульный (только # заголовок, без буллетов)
- Последний слайд — ключевые выводы

Формат вывода — чистый Markdown, ничего лишнего.
"""

transcripts = {}


def add_text_box(slide, text, left, top, width, height,
                 font_size=18, bold=False, color=(255, 255, 255),
                 align=PP_ALIGN.LEFT):
    """Добавляет текстовый блок на слайд с заданными параметрами."""
    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    tf = txBox.text_frame
    tf.word_wrap = True

    # Первый параграф
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(*color)
    return txBox


def build_pptx(md_content: str, output_path: str):
    prs = Presentation(TEMPLATE_PATH)

    # Слайд 10x5.62 inches (из шаблона)
    # Зоны контента: заголовок вверху слева, тело под ним
    TITLE_LEFT   = 1.38
    TITLE_TOP    = 0.16
    TITLE_WIDTH  = 7.50
    TITLE_HEIGHT = 0.45

    BODY_LEFT    = 1.38
    BODY_TOP     = 0.80
    BODY_WIDTH   = 7.50
    BODY_HEIGHT  = 4.50

    current_title = None
    current_bullets = []
    is_first_slide = True

    def flush_slide():
        nonlocal is_first_slide
        if current_title is None:
            return

        # Выбираем лейаут: 0=TITLE для первого слайда, 2=TITLE_AND_TWO_COLUMNS для остальных
        layout_idx = 0 if is_first_slide else 2
        layout = prs.slide_layouts[layout_idx]
        slide = prs.slides.add_slide(layout)
        is_first_slide = False

        # Пишем заголовок в placeholder idx=0
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == 0:
                ph.text = current_title
                for para in ph.text_frame.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(16)
                        run.font.bold = True
                break

        # Пишем буллеты как textbox в зоне тела
        if current_bullets:
            txBox = slide.shapes.add_textbox(
                Inches(BODY_LEFT), Inches(BODY_TOP),
                Inches(BODY_WIDTH), Inches(BODY_HEIGHT)
            )
            tf = txBox.text_frame
            tf.word_wrap = True

            for i, bullet in enumerate(current_bullets):
                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                p.alignment = PP_ALIGN.LEFT
                p.space_before = Pt(6)
                run = p.add_run()
                run.text = f"• {bullet}"
                run.font.size = Pt(14)
                run.font.color.rgb = RGBColor(255, 255, 255)

    for line in md_content.splitlines():
        line = line.strip()
        if line.startswith("# ") or line.startswith("## "):
            flush_slide()
            current_title = line.lstrip("#").strip()
            current_bullets = []
        elif line.startswith(("- ", "* ", "• ")):
            current_bullets.append(line[2:].strip())
        elif line and not line.startswith("#") and current_title and not current_bullets:
            # Подзаголовок на титульном слайде
            current_bullets.append(line)

    flush_slide()
    prs.save(output_path)


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
            {"role": "user", "content": transcripts[user_id]}
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
