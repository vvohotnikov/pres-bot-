# bot.py
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

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template.pptx")

bot = Bot(token=os.getenv("TELEGRAM_TOKEN"))
dp = Dispatcher()
client = AsyncOpenAI(
    api_key=os.getenv("OPENAI_API_KEY"),
    base_url="https://api.proxyapi.ru/openai/v1"
)

SYSTEM_PROMPT = """
Ты — ассистент для создания презентаций на русском языке.
Тебе дают расшифровку голосового сообщения.
Твоя задача — структурировать его в Markdown-презентацию.

Правила:
- Используй ВСЁ, что сказал пользователь. Разворачивай каждую мысль.
- Каждый слайд: # Заголовок + 4-6 буллетов. Каждый буллет — полное предложение 10-20 слов.
- Буллеты содержательны: раскрывай суть, используй детали из контекста.
- Минимум 5 слайдов, максимум 10.
- Первый слайд: только # заголовок (тема) и ## подзаголовок (одна фраза о чём речь).
- Последний слайд: # Ключевые выводы — 4-5 финальных тезисов.

Формат — строго чистый Markdown. Никаких пояснений, только слайды.
"""

transcripts = {}


def _clear_slides(prs: Presentation):
    xml_slides = prs.slides._sldIdLst
    for sld_id in list(xml_slides):
        xml_slides.remove(sld_id)


def build_pptx(md_content: str, output_path: str):
    prs = Presentation(TEMPLATE_PATH)
    layout_map = {l.name: l for l in prs.slide_layouts}

    # TITLE    — красивый титульный (большой фон, без мешающих картинок)
    # TITLE_ONLY — контентный, только фон + лого, нет перекрывающих shapes
    layout_title   = layout_map.get("TITLE")
    layout_content = layout_map.get("TITLE_ONLY")

    prs_out = Presentation(TEMPLATE_PATH)
    _clear_slides(prs_out)

    # Слайд 10.00 x 5.62 inches
    # Верхняя зона (лого + линия): 0 – 0.55"
    # Безопасная зона контента: от top=0.62"
    WHITE = (255, 255, 255)
    GRAY  = (180, 180, 180)

    current_title    = None
    current_subtitle = None
    current_bullets  = []
    is_first         = True

    def add_tb(slide, text, left, top, width, height,
               size=14, bold=False, color=WHITE, align=PP_ALIGN.LEFT, wrap=True):
        tb = slide.shapes.add_textbox(
            Inches(left), Inches(top), Inches(width), Inches(height)
        )
        tf = tb.text_frame
        tf.word_wrap = wrap
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = RGBColor(*color)
        return tb

    def flush_slide():
        nonlocal is_first, current_subtitle
        if current_title is None:
            return

        if is_first:
            slide = prs_out.slides.add_slide(layout_title)
            is_first = False
            # Титульный: заголовок крупно по центру слайда
            add_tb(slide, current_title,
                   left=0.40, top=1.60, width=9.20, height=1.10,
                   size=32, bold=True, color=WHITE)
            if current_subtitle:
                add_tb(slide, current_subtitle,
                       left=0.40, top=2.85, width=9.20, height=0.65,
                       size=16, bold=False, color=GRAY)
        else:
            slide = prs_out.slides.add_slide(layout_content)
            # Заголовок слайда
            add_tb(slide, current_title,
                   left=0.40, top=0.62, width=9.20, height=0.55,
                   size=20, bold=True, color=WHITE)
            # Буллеты
            if current_bullets:
                tb = slide.shapes.add_textbox(
                    Inches(0.40), Inches(1.30),
                    Inches(9.20), Inches(4.05)
                )
                tf = tb.text_frame
                tf.word_wrap = True
                for i, bullet in enumerate(current_bullets):
                    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                    p.alignment = PP_ALIGN.LEFT
                    p.space_before = Pt(6)
                    run = p.add_run()
                    run.text = f"— {bullet}"
                    run.font.size = Pt(13)
                    run.font.color.rgb = RGBColor(*WHITE)

        current_subtitle = None

    for line in md_content.splitlines():
        line = line.strip()
        if line.startswith("# "):
            flush_slide()
            current_title   = line[2:].strip()
            current_bullets = []
        elif line.startswith("## "):
            if is_first and current_title is not None and current_subtitle is None:
                current_subtitle = line[3:].strip()
            else:
                flush_slide()
                current_title   = line[3:].strip()
                current_bullets = []
        elif line.startswith(("- ", "* ", "• ")):
            current_bullets.append(line[2:].strip())
        elif line and current_title is not None and not line.startswith("#"):
            current_bullets.append(line)

    flush_slide()
    prs_out.save(output_path)


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
