# 🗺️ Google Maps Scraper & Automation Toolkit

A high-performance, asynchronous Google Maps data extraction toolkit built with **Python**, **Playwright**, and **Aiogram 3**. Extract business details, social media links, and customer reviews with export capabilities to CSV, Excel, JSON, and Word formats, accessible via an interactive CLI and a feature-rich Telegram Bot.

---

## ✨ Features

- ⚡ **Asynchronous Engine**: Built on top of `asyncio` and `playwright` for lightning-fast scrolling and concurrent detail extraction.
- 📍 **Places & Businesses Scraping**:
  - Name, Rating, Total Reviews, Category
  - Full Address, Latitude, Longitude
  - Phone Number, Official Website
  - Automatic Social Media Discovery (Instagram, Facebook, LinkedIn)
- ⭐ **Reviews Scraping**:
  - Extract detailed reviews (Reviewer name, rating, timestamp, and review text).
- 🤖 **Multi-Tier Telegram Bot**:
  - Interactive menus with Telegram inline keyboards.
  - Multi-tier quota management (`Free`, `Starter`, `Pro`, `Admin`) with SQLite persistence.
  - Live ASCII visual progress tracker with rate-limited message editing to prevent Telegram flood limits.
  - Admin management commands (set user tier, manage limits).
- 📊 **Multi-Format Export**:
  - **CSV** (Lightweight & clean)
  - **Excel (.xlsx)** (Formatted headers with clickable hyperlinks)
  - **JSON** (Full structured data)
  - **Word (.docx)** (Formatted report documents)
- 🛡️ **Anti-Detection & Proxy Support**:
  - Randomized User-Agents, custom delays, and optional HTTP/SOCKS proxy integration.

---

## 📁 Repository Structure

```text
├── modules/
│   ├── plugins/
│   │   ├── scraper.py           # Places & business scraper engine
│   │   └── review_scraper.py    # Reviews scraper engine
│   └── __init__.py
├── bot_telegram.py              # Aiogram 3 Telegram Bot service
├── main.py                      # Interactive CLI interface
├── config.ini.example           # Configuration template for Telegram bot
├── requirements.txt             # Project Python dependencies
├── start_bot.bat                # Windows helper script to start bot
├── stop_bot.bat                 # Windows helper script to stop bot
├── check_bot.ps1                # PowerShell status check utility
├── stop_bot_action.ps1          # PowerShell process termination utility
├── LICENSE                      # MIT License
└── README.md                    # Documentation
```

---

## 🚀 Getting Started

### Prerequisites
- **Python 3.10** or higher
- **Git**

### 1. Clone the Repository
```bash
git clone https://github.com/piscok123/Sraper-Maps.git
cd Sraper-Maps
```

### 2. Set Up a Virtual Environment
```bash
# Windows
python -m venv venv
venv\Scripts\activate

# Linux / macOS
python3 -m venv venv
source venv/bin/activate
```

### 3. Install Dependencies & Playwright Browser
```bash
pip install -r requirements.txt
playwright install chromium
```

### 4. Configuration (Telegram Bot)
If you plan to run the Telegram bot, copy `config.ini.example` to `config.ini`:
```bash
cp config.ini.example config.ini
```

Edit `config.ini` with your credentials:
```ini
[Telegram]
BotToken = YOUR_TELEGRAM_BOT_TOKEN_HERE

[Access]
AllowedUsers = YOUR_TELEGRAM_USER_ID
```
*(You can get a bot token from [@BotFather](https://t.me/BotFather) and your Telegram user ID from [@userinfobot](https://t.me/userinfobot)).*

---

## 💻 Usage

### Option A: Interactive CLI
Run the rich interactive terminal interface:
```bash
python main.py
```
Follow the interactive prompts to choose the scraper type (Places / Reviews), search query, limit, proxy settings, and export formats.

### Option B: Telegram Bot
Start the Telegram Bot service:
```bash
python bot_telegram.py
```
Or on Windows, simply double-click `start_bot.bat`.

#### Available Bot Commands:
- `/start` - Launch main menu and wizard
- `/places` - Start business places scraping wizard
- `/reviews` - Start reviews scraping wizard
- `/profile` - View your active tier, daily limits, and quota status
- `/help` - View usage guide
- `/settier <user_id> <tier> <days>` *(Admin only)* - Assign tier to a user

---

## ⚖️ Legal & Ethical Disclaimer

> **IMPORTANT**: This software is developed strictly for **educational, testing, and research purposes**. 
> 
> Automated data scraping may violate the **Terms of Service** of third-party platforms (including Google Maps). The authors and contributors of this repository do not endorse or encourage any unauthorized data harvesting. By using this tool, you agree that you are solely responsible for ensuring compliance with applicable local laws, regulations, and third-party terms of use.

---

## 📄 License

This project is licensed under the **MIT License** - see the [LICENSE](LICENSE) file for details.
