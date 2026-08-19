import os
import sys
import time
import asyncio
import questionary
from rich.console import Console
from rich.panel import Panel
from rich.text import Text

# Import Modules dari arsitektur baru
from modules.plugins.scraper import GoogleMapsScraperAsync
from modules.plugins.review_scraper import GoogleMapsReviewScraperAsync
from playwright.async_api import async_playwright

console = Console()

def format_duration(seconds: float) -> str:
    m, s = divmod(int(seconds), 60)
    h, m = divmod(m, 60)
    if h > 0:
        return f"{h} Jam {m} Menit {s} Detik"
    elif m > 0:
        return f"{m} Menit {s} Detik"
    else:
        return f"{s} Detik"

def display_welcome_banner():
    """Menampilkan banner UI pada CLI"""
    os.system('cls' if os.name == 'nt' else 'clear')
    title = Text("🗺️ Google Maps Scraper Toolkit", justify="center", style="bold cyan")
    console.print(Panel(title, border_style="cyan", subtitle="Modular Enterprise Edition"))
    print("\n")

def run_places_scraper():
    query = questionary.text("🔍 Masukkan kata kunci pencarian (cth: Hotel di Bandung):").ask()
    if not query: return
    
    limit_str = questionary.text("🎯 Maksimal data Tempat?", default="20").ask()
    if not limit_str: return
    limit = int(limit_str)
    
    headless = questionary.confirm("👻 Mode Headless (Sembunyikan UI Browser)?", default=True).ask()
    if headless is None: return

    # FITUR PROXY DIKEMBALIKAN
    use_proxy = questionary.confirm("🛡️ Gunakan Proxy (Opsional)?", default=False).ask()
    proxy_url = None
    if use_proxy:
        proxy_url = questionary.text("Masukkan Proxy URL (cth: http://user:pass@ip:port):").ask()
    
    formats = questionary.checkbox(
        "💾 Pilih format Export (Spasi untuk pilih, langsung Enter untuk default CSV):", 
        choices=["CSV", "JSON", "Excel (XLSX)", "Word (DOCX)"]
    ).ask()
    
    # Membedakan antara batal (Ctrl+C) dengan list kosong [] (lupa centang)
    if formats is None: 
        return
    
    console.print("\n[bold]🚀 Menjalankan Module: PLACES SCRAPER...[/bold]")
    # proxy_url dikirimkan ke modul
    scraper = GoogleMapsScraperAsync(query, limit, headless, proxy=proxy_url)
    
    start_time = time.time()
    asyncio.run(scraper.scrape())
    scraper.save_data(formats)
    end_time = time.time()
    
    duration_str = format_duration(end_time - start_time)
    console.print(f"[bold yellow]⏱️ Total waktu eksekusi: {duration_str}[/bold yellow]\n")

def run_reviews_scraper():
    url = questionary.text("🔗 Masukkan Tautan (URL) Google Maps Tempat tersebut:").ask()
    if not url: return
    
    limit_str = questionary.text("🎯 Maksimal data Ulasan yang ditarik?", default="50").ask()
    if not limit_str: return
    limit = int(limit_str)
    
    headless = questionary.confirm("👻 Mode Headless?", default=True).ask()
    if headless is None: return

    # FITUR PROXY DIKEMBALIKAN
    use_proxy = questionary.confirm("🛡️ Gunakan Proxy (Opsional)?", default=False).ask()
    proxy_url = None
    if use_proxy:
        proxy_url = questionary.text("Masukkan Proxy URL (cth: http://user:pass@ip:port):").ask()
    
    formats = questionary.checkbox(
        "💾 Pilih format Export (Spasi untuk pilih, langsung Enter untuk default CSV):", 
        choices=["CSV", "JSON", "Excel (XLSX)", "Word (DOCX)"]
    ).ask()
    
    # Membedakan antara batal (Ctrl+C) dengan list kosong [] (lupa centang)
    if formats is None: 
        return
    
    console.print("\n[bold]🚀 Menjalankan Module: REVIEWS SCRAPER...[/bold]")
    # proxy_url dikirimkan ke modul
    scraper = GoogleMapsReviewScraperAsync(url, limit, headless, proxy=proxy_url)
    
    start_time = time.time()
    asyncio.run(scraper.scrape())
    scraper.save_data(formats)
    end_time = time.time()
    
    duration_str = format_duration(end_time - start_time)
    console.print(f"[bold yellow]⏱️ Total waktu eksekusi: {duration_str}[/bold yellow]\n")

async def _login_async():
    session_dir = os.path.join(os.getcwd(), "browser_session")
    if not os.path.exists(session_dir):
        os.makedirs(session_dir)
        
    async with async_playwright() as p:
        console.print("[bold cyan]🤖 Menyiapkan Browser Profile (Persistent Context)...[/bold cyan]")
        context = await p.chromium.launch_persistent_context(
            user_data_dir=session_dir,
            headless=False,
            args=['--disable-blink-features=AutomationControlled']
        )
        page = context.pages[0] if len(context.pages) > 0 else await context.new_page()
        await page.goto("https://accounts.google.com/")
        
        console.print("\n[bold green]✅ Browser terbuka![/bold green]")
        console.print("[yellow]Silakan login ke akun Google Dummy Anda di jendela browser yang terbuka.[/yellow]")
        console.print("[yellow]Setelah Anda selesai login dan semuanya beres, kembali ke sini.[/yellow]\n")
        
        # Menggunakan loop tidur sederhana sambil menunggu input (asyncio workaround)
        await asyncio.to_thread(input, "Tekan ENTER di sini JIKA ANDA SUDAH SELESAI LOGIN untuk menyimpan sesi... ")
        
        console.print("[bold green]💾 Menyimpan sesi dan menutup browser...[/bold green]")
        await context.close()

def run_login_mode():
    console.print("\n[bold]🚀 Menjalankan Module: SETUP LOGIN AKUN...[/bold]")
    asyncio.run(_login_async())

def main() -> None:
    display_welcome_banner()
    
    mode = questionary.select(
        "🛠️ Pilih Alat Scraping yang ingin digunakan:",
        choices=[
            "1. Scraper Lokasi & Tempat (Places)",
            "2. Scraper Ulasan (Reviews)",
            "3. Setup Login Akun Google (Browser Session)",
            "4. Keluar"
        ]
    ).ask()

    if mode == "1. Scraper Lokasi & Tempat (Places)":
        run_places_scraper()
    elif mode == "2. Scraper Ulasan (Reviews)":
        run_reviews_scraper()
    elif mode == "3. Setup Login Akun Google (Browser Session)":
        run_login_mode()
    else:
        console.print("[yellow]Sampai jumpa![/yellow]")
        sys.exit(0)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        console.print("\n[bold red]⛔ Eksekusi dihentikan manual oleh pengguna.[/bold red]")
        sys.exit(1)