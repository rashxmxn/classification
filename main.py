import os
import pandas as pd
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from dotenv import load_dotenv
import logging

load_dotenv()


# Enable logging
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

class EducationalProgramBot:
    def __init__(self, excel_file_path):
        """
        Initialize the bot with Excel file containing educational programs

        Excel structure expected:
        Column A: Code of Educational Program (EP)
        Column B: Name of Educational Program
        Column C: Cipher codes
        Column D: Name of cipher
        """
        self.excel_file_path = excel_file_path
        self.data = None
        self.load_data()

    def load_data(self):
        """Load data from Excel file"""
        try:
            # Read Excel file
            self.data = pd.read_excel(self.excel_file_path, header=None)

            # Check if we have at least 3 columns
            if len(self.data.columns) < 3:
                logger.error(f"Excel file has only {len(self.data.columns)} columns, expected at least 3")
                self.data = pd.DataFrame()
                return

            # Rename columns for easier access (adjust based on actual structure)
            column_names = ['ep_code', 'ep_name', 'cipher_code']
            if len(self.data.columns) >= 4:
                column_names.append('cipher_name')

            # Only rename up to the number of columns we actually have
            self.data.columns = column_names[:len(self.data.columns)]

            # Process the data to handle the specific structure
            processed_data = []
            current_ep_code = None
            current_ep_name = None

            for index, row in self.data.iterrows():
                # Check if this row has EP code and name (main educational program row)
                if pd.notna(row['ep_code']) and pd.notna(row['ep_name']):
                    current_ep_code = str(row['ep_code']).strip()
                    current_ep_name = str(row['ep_name']).strip()

                    # If this row also has a cipher code, add it
                    if pd.notna(row['cipher_code']):
                        cipher_code = str(row['cipher_code']).strip()
                        if cipher_code not in ['nan', '']:
                            processed_data.append({
                                'ep_code': current_ep_code,
                                'ep_name': current_ep_name,
                                'cipher_code': cipher_code
                            })

                # If this row only has cipher code (continuation of previous EP)
                elif pd.notna(row['cipher_code']) and current_ep_code and current_ep_name:
                    cipher_code = str(row['cipher_code']).strip()
                    if cipher_code not in ['nan', '']:
                        processed_data.append({
                            'ep_code': current_ep_code,
                            'ep_name': current_ep_name,
                            'cipher_code': cipher_code
                        })

            # Create new DataFrame from processed data
            self.data = pd.DataFrame(processed_data)

            # Show some sample cipher codes
            if len(self.data) > 0:
                sample_ciphers = self.data['cipher_code'].head(10).tolist()
                logger.info(f"Sample cipher codes: {sample_ciphers}")

            logger.info(f"Successfully loaded {len(self.data)} records from Excel file")

        except Exception as e:
            logger.error(f"Error loading Excel file: {e}")
            self.data = pd.DataFrame()

    def search_programs(self, cipher_query):
        """
        Search for educational programs matching the cipher query

        Args:
            cipher_query (str): The cipher code to search for (e.g., "070107 3")

        Returns:
            list: List of unique matching programs with format "EP_Code - EP_Name"
        """
        if self.data is None or self.data.empty:
            return []

        # Clean the query
        cipher_query = cipher_query.strip()

        # Search for exact matches first
        exact_matches = self.data[self.data['cipher_code'] == cipher_query]

        if not exact_matches.empty:
            results = []
            seen = set()  # To track duplicates

            for _, row in exact_matches.iterrows():
                result = f"{row['ep_code']} - {row['ep_name']}"
                # Only add if we haven't seen this exact combination before
                if result not in seen:
                    results.append(result)
                    seen.add(result)
            return results

        # If no exact match, try partial matching
        partial_matches = self.data[self.data['cipher_code'].str.contains(cipher_query, case=False, na=False)]

        if not partial_matches.empty:
            results = []
            seen = set()  # To track duplicates

            for _, row in partial_matches.iterrows():
                result = f"{row['ep_code']} - {row['ep_name']}"
                # Only add if we haven't seen this exact combination before
                if result not in seen:
                    results.append(result)
                    seen.add(result)
            return results

        return []

# Initialize bot instance (will be set when Excel file is provided)
bot_instance = None

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Send a message when the command /start is issued."""
    welcome_message = """
🎓 Добро пожаловать в бот поиска образовательных программ!

Отправьте мне шифр специальности (например: "4S03220203"), и я найду все подходящие образовательные программы.

Команды:
/start - Показать это сообщение
/help - Помощь
/status - Проверить статус бота

Просто отправьте шифр специальности для поиска!
    """
    await update.message.reply_text(welcome_message)

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Send a message when the command /help is issued."""
    help_text = """
📚 Как использовать данного бота:

1. Отправьте шифр специальности
2. Бот найдет все соответствующие образовательные программы
3. Бот сформулирует ответ для абитуриента и отправит его в следующем формате: Код программы - Название программы + Перечень документов

Примеры запросов:
• 070107 3
• 4S03220203
• 5AB02140101

❗ Убедитесь, что шифр написан правильно!
    """
    await update.message.reply_text(help_text)

async def status_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Check bot status and data availability."""
    global bot_instance

    if bot_instance is None or bot_instance.data is None or bot_instance.data.empty:
        status_text = "❌ Данные не загружены. Обратитесь к администратору."
    else:
        record_count = len(bot_instance.data)
        status_text = f"✅ Бот работает нормально\n📊 Загружено записей: {record_count}"

    await update.message.reply_text(status_text)

async def search_cipher(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle cipher search requests."""
    global bot_instance

    if bot_instance is None:
        await update.message.reply_text("❌ Бот не настроен. Обратитесь к администратору.")
        return

    cipher_query = update.message.text.strip()

    # Search for matching programs
    results = bot_instance.search_programs(cipher_query)

    if results:
        # Format header normally
        response = f"Найденные программы для шифра '{cipher_query}':\n\n"

        # Format results in monospace for easy copying
        monospace_results = "\n".join(results)
        response += f"```\nВы можете поступить в наш университет по данным ГОП:\n{monospace_results}\n\nСписок необходимых документов:\n1. Диплом с приложением (оригинал + копия)\n2. Сертификат ЕНТ при сдаче.\n3. Копия удостоверения личности.\n4. Медицинская справка формы 075-у со снимком флюрографии.\n5. Медицинская справка формы 063 (паспорт здоровья).\n6. Фотография 3х4 в электронном формате.\n\nДля дополнительной информации: https://www.ektu.kz/admissiondetails.aspx?ttab=1```"

        # Add count information normally
        response += f"\n\nВсего найдено: {len(results)} программ(ы)"

    else:
        response = f"❌ Не найдено программ для шифра '{cipher_query}'\n\n"
        response += "Проверьте правильность написания шифра.\n\n Если данная ошибка повторяется, то у нас нет ГОП по данному шифру"

    await update.message.reply_text(response, parse_mode='Markdown')

async def error_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Log the error and send a telegram message to notify the developer."""
    logger.warning('Update "%s" caused error "%s"', update, context.error)

    if update and update.message:
        await update.message.reply_text(
            "😔 Произошла ошибка при обработке вашего запроса. "
            "Пожалуйста, попробуйте еще раз или обратитесь к администратору."
        )

def main():
    """Start the bot."""
    global bot_instance

    # Configuration
    TOKEN = os.getenv("BOT_TOKEN")
    EXCEL_FILE_PATH = "./educational_programs.xlsx"

    # Initialize the educational program bot
    try:
        bot_instance = EducationalProgramBot(EXCEL_FILE_PATH)
        logger.info("Educational program bot initialized successfully")
    except Exception as e:
        logger.error(f"Failed to initialize bot: {e}")
        return

    # Create the Application
    application = Application.builder().token(TOKEN).build()

    # Add handlers
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler("status", status_command))

    # Handle all text messages as cipher search queries
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, search_cipher))

    # Add error handler
    application.add_error_handler(error_handler)

    # Run the bot until the user presses Ctrl-C
    logger.info("Starting bot...")
    application.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == '__main__':
    main()