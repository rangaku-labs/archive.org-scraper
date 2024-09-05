import os
import requests
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from concurrent.futures import ThreadPoolExecutor, as_completed
from functools import lru_cache
import logging
import threading
import time
import re
import traceback
import shelve
import hashlib
import json
import pickle
import multiprocessing

# Import the AutocompleteCombobox class
from ttkwidgets.autocomplete import AutocompleteCombobox

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

@lru_cache(maxsize=32)
def fetch_file_data(item_url, file_types):
    result = []
    try:
        response = requests.get(item_url)
        item_soup = BeautifulSoup(response.content, 'html.parser')
        file_links = set()
        for file_type in file_types:
            download_links = item_soup.select(f'a.download-pill[href$=".{file_type}"]')
            for link in download_links:
                download_url = f"https://archive.org{link['href']}"
                if download_url not in file_links:
                    file_links.add(download_url)
                    file_name = os.path.basename(download_url)
                    file_size = get_file_size(link, item_soup)
                    book_name = get_book_name(item_soup)
                    description = get_book_description(item_soup)
                    result.append((file_name, book_name, download_url, file_size, description))
    except Exception as e:
        logging.error(f"Error fetching data for {item_url}: {e}")
    return result

def get_file_size(link, soup):
    # Try to get size from data-original-title
    size = link.get('data-original-title')
    if size:
        return size

    # Try to get size from title attribute
    size = link.get('title')
    if size:
        return size

    # Try to find size in the page content
    size_elem = soup.select_one('.item-stats .size')
    if size_elem:
        return size_elem.text.strip()

    return 'Unknown'

def get_book_name(item_soup):
    title_tag = item_soup.find('h1', class_='item-title')
    return title_tag.text.strip() if title_tag else "Unknown"

def get_book_description(item_soup):
    description_elem = item_soup.select_one('div[itemprop="description"]')
    return description_elem.text.strip() if description_elem else "No description available."

# Global variables
all_items = []
fetch_thread = None
search_history = []
common_file_types = ['pdf', 'epub', 'mobi', 'txt', 'doc', 'docx', 'rtf', 'djvu']

class SearchCache:
    def __init__(self, cache_file='search_cache'):
        self.cache_file = cache_file

    def get_cache_key(self, search_params, file_types):
        key_data = json.dumps({**search_params, 'file_types': file_types}, sort_keys=True).encode('utf-8')
        return hashlib.md5(key_data).hexdigest()

    def get(self, search_params, file_types):
        key = self.get_cache_key(search_params, file_types)
        with shelve.open(self.cache_file) as cache:
            if key in cache:
                return cache[key]['data']
        return None

    def set(self, search_params, file_types, data):
        key = self.get_cache_key(search_params, file_types)
        with shelve.open(self.cache_file) as cache:
            cache[key] = {'data': data}

search_cache = SearchCache()

class LazyTreeview(ttk.Treeview):
    def __init__(self, master, **kw):
        ttk.Treeview.__init__(self, master, **kw)
        self._items = []
        self.bind("<<TreeviewOpen>>", self._on_open)

    def set_items(self, items):
        self._items = items
        self.delete(*self.get_children())
        if len(items) > 0:
            self.insert("", "end", "dummy")

    def _on_open(self, event):
        self.delete("dummy")
        for i, item in enumerate(self._items):
            self.insert("", "end", values=item[:4], iid=str(i))

class FetchFilesThread(threading.Thread):
    def __init__(self, base_url, file_types, search_params):
        threading.Thread.__init__(self)
        self.base_url = base_url
        self.file_types = file_types
        self.search_params = search_params
        self.total_files = 0
        self.total_size = 0
        self.is_paused = False
        self.is_cancelled = False
        self.total_available_files = 0
        self.progress = 0
        self.page = 1

    def run(self):
        global all_items
        all_items = []
        
        while self.total_files < self.total_available_files or self.total_available_files == 0:
            if self.is_cancelled:
                break
            
            url = f"{self.base_url}&page={self.page}"
            response = requests.get(url)
            data = response.json()
            
            if "response" in data and "docs" in data["response"]:
                docs = data["response"]["docs"]
                self.total_available_files = int(data["response"]["numFound"])
                self.update_status()
                
                if not docs:  # If we've reached a page with no results, break the loop
                    break
                
                with ThreadPoolExecutor(max_workers=10) as executor:
                    item_futures = []
                    for doc in docs:
                        identifier = doc["identifier"]
                        item_url = f"https://archive.org/details/{identifier}"
                        item_futures.append(executor.submit(fetch_file_data, item_url, tuple(self.file_types)))
                    
                    for item_future in as_completed(item_futures):
                        if self.is_cancelled:
                            break
                        while self.is_paused:
                            time.sleep(0.1)
                        result = item_future.result()
                        for file_name, book_name, file_url, file_size, description in result:
                            parsed_size = parse_size(file_size)
                            formatted_size = format_size(parsed_size)
                            self.total_files += 1
                            self.total_size += parsed_size
                            self.progress = min((self.total_files / self.total_available_files) * 100, 100)
                            item_id = file_tree.insert("", "end", values=(file_name, book_name, formatted_size, file_url))
                            all_items.append((file_name, book_name, formatted_size, file_url, description, item_id))
                            self.update_status()
            
            self.page += 1  # Move to the next page

        if not self.is_cancelled:
            status_label.config(text=f"Fetching complete. Found {self.total_files} files out of {self.total_available_files}")
            total_size_label.config(text=f"Total size: {format_size(self.total_size)}")
        progress_bar["value"] = 100
        window.update()

    def update_status(self):
        status_label.config(text=f"Found {self.total_files} files out of {self.total_available_files}")
        total_size_label.config(text=f"Total size: {format_size(self.total_size)}")
        progress_bar["value"] = self.progress

def create_icon():
    icon = tk.PhotoImage(width=64, height=64)
    icon.put(("green",), to=(5, 5, 59, 59))  # Green background
    icon.put(("black",), to=(10, 10, 54, 54))  # Black inner square
    icon.put(("green",), to=(20, 20, 44, 44))  # Green inner square
    icon.put(("black",), to=(25, 25, 39, 39))  # Black center
    return icon

def show_splash_screen():
    splash = tk.Toplevel()
    splash.overrideredirect(True)
    icon = create_icon()
    icon = icon.zoom(5)  # Make the icon larger for the splash screen
    label = tk.Label(splash, image=icon, bg='#001100')
    label.image = icon  # Keep a reference to prevent garbage collection
    label.pack()
    splash.geometry(f"320x320+{(splash.winfo_screenwidth()-320)//2}+{(splash.winfo_screenheight()-320)//2}")
    splash.after(3000, splash.destroy)
    return splash

def build_advanced_query(**kwargs):
    query_parts = []
    if kwargs.get('language'):
        query_parts.append(f"language:{kwargs['language']}")
    if kwargs.get('start_year') and kwargs.get('end_year'):
        query_parts.append(f"year:[{kwargs['start_year']} TO {kwargs['end_year']}]")
    if kwargs.get('keyword'):
        query_parts.append(f"title:{kwargs['keyword']}")
    if kwargs.get('author'):
        query_parts.append(f"creator:{kwargs['author']}")
    return " AND ".join(query_parts)

def perform_search():
    global fetch_thread
    search_params = {
        'language': language_entry.get().strip(),
        'start_year': start_year_entry.get().strip(),
        'end_year': end_year_entry.get().strip(),
        'keyword': keyword_entry.get().strip(),
        'author': author_entry.get().strip()
    }
    file_types = [ft.strip() for ft in file_type_entry.get().split(',') if ft.strip()]
    
    # Add to search history
    search_history.append({**search_params, 'file_types': file_types})
    if len(search_history) > 10:
        search_history.pop(0)
    
    # Check cache
    cached_results = search_cache.get(search_params, file_types)
    if cached_results:
        display_cached_results(cached_results)
        return

    logging.debug(f"Search params: {search_params}")
    logging.debug(f"File types: {file_types}")
    
    # Remove empty parameters
    search_params = {k: v for k, v in search_params.items() if v}
    
    if not search_params:
        messagebox.showerror("Error", "Please provide at least one search parameter.")
        return
    
    query = build_advanced_query(**search_params)
    logging.debug(f"Built query: {query}")
    
    base_url = f"https://archive.org/advancedsearch.php?q={query}&fl[]=identifier&fl[]=title&fl[]=creator&fl[]=year&fl[]=subject&fl[]=description&sort[]=downloads+desc&rows=100&output=json"
    logging.debug(f"Base Search URL: {base_url}")
    
    # Clear existing results
    file_tree.delete(*file_tree.get_children())
    status_label.config(text="Searching...")
    total_size_label.config(text="Total size: 0 B")
    progress_bar["value"] = 0
    
    # Start a new thread to fetch results
    fetch_thread = FetchFilesThread(base_url, file_types, search_params)
    fetch_thread.start()

    # Enable pause button and disable search button
    pause_button.config(state=tk.NORMAL)
    search_button.config(state=tk.DISABLED)
    
    logging.debug("Search initiated successfully")

def display_cached_results(results):
    global all_items
    all_items = results
    file_tree.delete(*file_tree.get_children())
    total_size = 0
    for item in all_items:
        file_tree.insert("", "end", iid=item[-1], values=item[:4])  # Exclude description from display
        total_size += parse_size(item[2])
    status_label.config(text=f"Loaded {len(results)} files from cache")
    total_size_label.config(text=f"Total size: {format_size(total_size)}")
    progress_bar["value"] = 100

def pause_resume_fetch():
    global fetch_thread
    if fetch_thread and fetch_thread.is_alive():
        if fetch_thread.is_paused:
            fetch_thread.is_paused = False
            pause_button.config(text="Pause")
            status_label.config(text="Resuming...")
        else:
            fetch_thread.is_paused = True
            pause_button.config(text="Resume")
            status_label.config(text="Paused")
    window.update()

def confirm_cancel():
    if messagebox.askyesno("Confirm Cancel", "Are you sure you want to cancel the current search?"):
        cancel_fetch()

def cancel_fetch():
    global fetch_thread
    if fetch_thread and fetch_thread.is_alive():
        fetch_thread.is_cancelled = True
        status_label.config(text="Cancelling...")
        window.update()
        fetch_thread.join(timeout=5)  # Wait for up to 5 seconds
        if fetch_thread.is_alive():
            logging.warning("Thread did not terminate within timeout.")
    clear_gui()
    search_button.config(state=tk.NORMAL)
    pause_button.config(state=tk.DISABLED)
    status_label.config(text="Search cancelled.")
    fetch_thread = None

def handle_download_error(file_name, error):
    log_message(f"Error downloading {file_name}: {error}")
    retry = messagebox.askretrycancel("Download Error", f"Error downloading {file_name}. Retry?")
    if retry:
        download_selected_files_thread()

def parse_size(size_str):
    if not size_str or size_str.lower() == 'unknown':
        return 0
    size_str = size_str.replace(',', '').strip()
    match = re.match(r"([\d.]+)\s*([KMGT]?B?)", size_str, re.I)
    if not match:
        return 0
    size, unit = match.groups()
    size = float(size)
    unit = unit.upper() if unit else 'B'
    units = {'B': 1, 'K': 1024, 'M': 1024**2, 'G': 1024**3, 'T': 1024**4}
    return int(size * units.get(unit[:1], 1))

def format_size(size):
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if size < 1024.0:
            return f"{size:.2f} {unit}"
        size /= 1024.0

def download_selected_files():
    thread = threading.Thread(target=download_selected_files_thread)
    thread.start()

def download_selected_files_thread():
    selected_items = file_tree.selection()
    download_dir = download_dir_entry.get()

    if not download_dir:
        messagebox.showerror("Error", "Please select a download directory.")
        return

    total_files = len(selected_items)
    downloaded_files = 0

    for item in selected_items:
        download_url = file_tree.item(item, "values")[3]
        file_name = file_tree.item(item, "values")[0]
        file_path = os.path.join(download_dir, file_name)

        try:
            status_label.config(text=f"Downloading: {file_name} ({downloaded_files+1}/{total_files})")
            progress_bar["value"] = (downloaded_files / total_files) * 100
            window.update()

            response = requests.get(download_url, stream=True)
            
            if response.status_code == 200:
                with open(file_path, "wb") as file:
                    for chunk in response.iter_content(chunk_size=8192):
                        file.write(chunk)
            else:
                status_label.config(text=f"Skipping: {file_name} (status code: {response.status_code})")
                window.update()
                continue
            
            downloaded_files += 1
            window.update()
        except Exception as e:
            status_label.config(text=f"Error downloading: {file_name}")
            window.update()

    status_label.config(text="Download complete.")
    progress_bar["value"] = 100
    window.update()

def select_download_dir():
    download_dir = filedialog.askdirectory()
    download_dir_entry.delete(0, tk.END)
    download_dir_entry.insert(0, download_dir)

def clear_gui():
    file_tree.delete(*file_tree.get_children())
    status_label.config(text="")
    total_size_label.config(text="")
    progress_bar["value"] = 0

def export_results():
    if not all_items:
        show_error("Export Error", "No results to export.")
        return
    
    file_types = [("Text files", "*.txt"), ("CSV files", "*.csv"), ("JSON files", "*.json")]
    file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=file_types)
    if not file_path:
        return
    
    try:
        if file_path.endswith('.txt'):
            export_as_txt(file_path)
        elif file_path.endswith('.csv'):
            export_as_csv(file_path)
        elif file_path.endswith('.json'):
            export_as_json(file_path)
        messagebox.showinfo("Export Successful", f"Results exported to {file_path}")
    except Exception as e:
        show_error("Export Error", f"Failed to export results: {str(e)}")

def export_as_txt(file_path):
    with open(file_path, 'w', encoding='utf-8') as f:
        for item in all_items:
            f.write(f"File Name: {item[0]}\n")
            f.write(f"Book Name: {item[1]}\n")
            f.write(f"Size: {item[2]}\n")
            f.write(f"URL: {item[3]}\n")
            f.write(f"Description: {item[4]}\n")
            f.write("\n---\n\n")

def export_as_csv(file_path):
    with open(file_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['File Name', 'Book Name', 'Size', 'URL', 'Description'])
        for item in all_items:
            writer.writerow(item[:5])  # Exclude the item_id

def export_as_json(file_path):
    data = [{'File Name': item[0], 'Book Name': item[1], 'Size': item[2], 'URL': item[3], 'Description': item[4]} for item in all_items]
    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def show_error(title, message):
    messagebox.showerror(title, message)

def show_detailed_view(event):
    selected_item = file_tree.selection()[0]
    item_data = file_tree.item(selected_item)['values']
    description = next((item[4] for item in all_items if item[:4] == item_data), "No description available.")
    
    detail_window = tk.Toplevel(window)
    detail_window.title("Item Details")
    detail_window.geometry("600x400")
    detail_window.configure(bg='#001100')

    ttk.Label(detail_window, text="File Name:", style="TLabel").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    ttk.Label(detail_window, text=item_data[0], style="TLabel").grid(row=0, column=1, sticky="w", padx=5, pady=5)

    ttk.Label(detail_window, text="Book Name:", style="TLabel").grid(row=1, column=0, sticky="w", padx=5, pady=5)
    ttk.Label(detail_window, text=item_data[1], style="TLabel").grid(row=1, column=1, sticky="w", padx=5, pady=5)

    ttk.Label(detail_window, text="Size:", style="TLabel").grid(row=2, column=0, sticky="w", padx=5, pady=5)
    ttk.Label(detail_window, text=item_data[2], style="TLabel").grid(row=2, column=1, sticky="w", padx=5, pady=5)

    ttk.Label(detail_window, text="URL:", style="TLabel").grid(row=3, column=0, sticky="w", padx=5, pady=5)
    ttk.Label(detail_window, text=item_data[3], style="TLabel").grid(row=3, column=1, sticky="w", padx=5, pady=5)

    ttk.Label(detail_window, text="Description:", style="TLabel").grid(row=4, column=0, sticky="nw", padx=5, pady=5)
    description_text = tk.Text(detail_window, wrap=tk.WORD, width=50, height=10, bg='#003300', fg='#00FF00')
    description_text.grid(row=4, column=1, sticky="nsew", padx=5, pady=5)
    description_text.insert(tk.END, description)
    description_text.config(state=tk.DISABLED)

    detail_window.grid_columnconfigure(1, weight=1)
    detail_window.grid_rowconfigure(4, weight=1)

def save_search_history():
    with open('search_history.pkl', 'wb') as f:
        pickle.dump(search_history, f)

def load_search_history():
    global search_history
    try:
        with open('search_history.pkl', 'rb') as f:
            search_history = pickle.load(f)
    except FileNotFoundError:
        search_history = []

def add_to_search_history(search_params):
    global search_history
    search_history.append(search_params)
    if len(search_history) > 10:  # Keep only the last 10 searches
        search_history.pop(0)
    save_search_history()

def show_search_history():
    history_window = tk.Toplevel(window)
    history_window.title("Search History")
    history_window.geometry("400x300")
    history_window.configure(bg='#001100')

    history_list = tk.Listbox(history_window, bg='#003300', fg='#00FF00', selectbackground='#005500')
    history_list.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    for search in reversed(search_history):
        history_text = f"{search.get('keyword', 'N/A')} ({search.get('language', 'N/A')}, {search.get('start_year', 'N/A')}-{search.get('end_year', 'N/A')})"
        history_list.insert(tk.END, history_text)

    def use_selected_search():
        selected_index = history_list.curselection()[0]
        selected_search = search_history[-(selected_index+1)]
        language_entry.delete(0, tk.END)
        language_entry.insert(0, selected_search.get('language', ''))
        start_year_entry.delete(0, tk.END)
        start_year_entry.insert(0, selected_search.get('start_year', ''))
        end_year_entry.delete(0, tk.END)
        end_year_entry.insert(0, selected_search.get('end_year', ''))
        keyword_entry.delete(0, tk.END)
        keyword_entry.insert(0, selected_search.get('keyword', ''))
        author_entry.delete(0, tk.END)
        author_entry.insert(0, selected_search.get('author', ''))
        file_type_entry.delete(0, tk.END)
        file_type_entry.insert(0, ', '.join(selected_search.get('file_types', [])))
        history_window.destroy()

    use_button = ttk.Button(history_window, text="Use Selected", command=use_selected_search)
    use_button.pack(pady=10)

def save_preferences():
    preferences = {
        'download_dir': download_dir_entry.get(),
        'language': language_entry.get(),
        'start_year': start_year_entry.get(),
        'end_year': end_year_entry.get(),
        'keyword': keyword_entry.get(),
        'author': author_entry.get(),
        'file_types': file_type_entry.get()
    }
    with open('user_preferences.json', 'w') as f:
        json.dump(preferences, f)

def load_preferences():
    try:
        with open('user_preferences.json', 'r') as f:
            preferences = json.load(f)
        download_dir_entry.insert(0, preferences.get('download_dir', ''))
        language_entry.insert(0, preferences.get('language', ''))
        start_year_entry.insert(0, preferences.get('start_year', ''))
        end_year_entry.insert(0, preferences.get('end_year', ''))
        keyword_entry.insert(0, preferences.get('keyword', ''))
        author_entry.insert(0, preferences.get('author', ''))
        file_type_entry.set(preferences.get('file_types', 'pdf'))
    except FileNotFoundError:
        pass

def create_gui():
    global window, file_tree, download_dir_entry, status_label, total_size_label, progress_bar, search_button, pause_button, language_entry, start_year_entry, end_year_entry, keyword_entry, author_entry, file_type_entry

    window = tk.Tk()
    window.title("ARCHIVE.ORG SCRAPER")
    window.geometry("1200x800")
    window.configure(bg='#001100')
    window.grid_columnconfigure(0, weight=1)
    window.grid_rowconfigure(1, weight=1)

    style = ttk.Style()
    style.theme_use('clam')
    style.configure("TFrame", background="#001100")
    style.configure("TLabel", background="#001100", foreground="#00FF00", font=("Courier", 10))
    style.configure("TEntry", fieldbackground="#003300", foreground="#00FF00", font=("Courier", 10))
    style.configure("TButton", background="#005500", foreground="#00FF00", font=("Courier", 10))
    style.configure("Treeview", background="#001100", foreground="#00FF00", fieldbackground="#003300", font=("Courier", 9))
    style.configure("Treeview.Heading", background="#005500", foreground="#00FF00", font=("Courier", 10, "bold"))
    style.configure("TProgressbar", troughcolor="#003300", background="#00FF00", thickness=20)

    # Search Parameters
    search_frame = ttk.Frame(window, padding=10)
    search_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
    search_frame.grid_columnconfigure((1, 3, 5), weight=1)

    ttk.Label(search_frame, text="Language:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    language_entry = ttk.Entry(search_frame)
    language_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)

    ttk.Label(search_frame, text="Start Year:").grid(row=0, column=2, sticky="w", padx=5, pady=5)
    start_year_entry = ttk.Entry(search_frame, width=10)
    start_year_entry.grid(row=0, column=3, sticky="w", padx=5, pady=5)

    ttk.Label(search_frame, text="End Year:").grid(row=0, column=4, sticky="w", padx=5, pady=5)
    end_year_entry = ttk.Entry(search_frame, width=10)
    end_year_entry.grid(row=0, column=5, sticky="w", padx=5, pady=5)

    ttk.Label(search_frame, text="Keyword:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
    keyword_entry = ttk.Entry(search_frame)
    keyword_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)

    ttk.Label(search_frame, text="Author:").grid(row=1, column=2, sticky="w", padx=5, pady=5)
    author_entry = ttk.Entry(search_frame)
    author_entry.grid(row=1, column=3, columnspan=3, sticky="ew", padx=5, pady=5)

    ttk.Label(search_frame, text="File Types:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
    file_type_entry = AutocompleteCombobox(search_frame, completevalues=common_file_types)
    file_type_entry.grid(row=2, column=1, columnspan=5, sticky="ew", padx=5, pady=5)
    file_type_entry.set('pdf')  # Set default value to 'pdf'

    search_button = ttk.Button(search_frame, text="Search", command=perform_search)
    search_button.grid(row=3, column=0, columnspan=2, pady=10)

    pause_button = ttk.Button(search_frame, text="Pause", command=pause_resume_fetch, state=tk.DISABLED)
    pause_button.grid(row=3, column=2, columnspan=2, pady=10)

    cancel_button = ttk.Button(search_frame, text="Cancel", command=confirm_cancel)
    cancel_button.grid(row=3, column=4, columnspan=2, pady=10)

    history_button = ttk.Button(search_frame, text="History", command=show_search_history)
    history_button.grid(row=3, column=6, pady=10)

    # Results
    results_frame = ttk.Frame(window, padding=10)
    results_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
    results_frame.grid_columnconfigure(0, weight=1)
    results_frame.grid_rowconfigure(0, weight=1)

    tree_frame = ttk.Frame(results_frame)
    tree_frame.grid(row=0, column=0, sticky="nsew")
    tree_frame.grid_columnconfigure(0, weight=1)
    tree_frame.grid_rowconfigure(0, weight=1)

    file_tree = LazyTreeview(tree_frame, columns=("name", "book_name", "size", "url"), show="headings", selectmode="extended")
    file_tree.heading("name", text="File Name", command=lambda: sort_tree("name", False))
    file_tree.heading("book_name", text="Book Name", command=lambda: sort_tree("book_name", False))
    file_tree.heading("size", text="Size", command=lambda: sort_tree("size", False))
    file_tree.heading("url", text="URL", command=lambda: sort_tree("url", False))
    file_tree.column("name", width=200, minwidth=200)
    file_tree.column("book_name", width=300, minwidth=300)
    file_tree.column("size", width=100, minwidth=100)
    file_tree.column("url", width=400, minwidth=400)
    file_tree.grid(row=0, column=0, sticky="nsew")

    tree_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=file_tree.yview)
    tree_scrollbar.grid(row=0, column=1, sticky="ns")

    file_tree.configure(yscrollcommand=tree_scrollbar.set)
    file_tree.bind("<Double-1>", show_detailed_view)

    # Download
    download_frame = ttk.Frame(window, padding=10)
    download_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=10)
    download_frame.grid_columnconfigure(1, weight=1)

    ttk.Label(download_frame, text="Download Directory:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    download_dir_entry = ttk.Entry(download_frame)
    download_dir_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
    
    select_dir_button = ttk.Button(download_frame, text="Select", command=select_download_dir)
    select_dir_button.grid(row=0, column=2, padx=5, pady=5)

    download_button = ttk.Button(download_frame, text="Download Selected", command=download_selected_files)
    download_button.grid(row=1, column=0, pady=10, padx=5)

    export_button = ttk.Button(download_frame, text="Export Results", command=export_results)
    export_button.grid(row=1, column=1, pady=10, padx=5)

    # Status bar
    status_frame = ttk.Frame(window)
    status_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=5)
    status_frame.grid_columnconfigure(0, weight=1)

    status_label = ttk.Label(status_frame, text="")
    status_label.grid(row=0, column=0, sticky="w")

    total_size_label = ttk.Label(status_frame, text="")
    total_size_label.grid(row=0, column=1, sticky="e")

    progress_bar = ttk.Progressbar(window, length=400, mode="determinate", style="TProgressbar")
    progress_bar.grid(row=4, column=0, sticky="ew", padx=10, pady=5)

    # Attribution label
    attribution_label = ttk.Label(window, text="CREATED BY RANGAKU RESEARCH LABS, A SUBSIDIARY OF RANGAKU ZAIBATSU INTERNATIONAL TRADING COMPANY. ALL BASES BELONGED TO US. ", foreground="#00FF00", background="#001100", font=("Courier", 8))
    attribution_label.grid(row=5, column=0, sticky="ew", padx=10, pady=5)

    # Set icon and show splash screen
    window.iconphoto(True, create_icon())
    splash = show_splash_screen()
    window.withdraw()
    window.after(3000, window.deiconify)

    # Load preferences
    load_preferences()

    return (window, file_tree, download_dir_entry, 
            status_label, total_size_label, progress_bar, search_button, 
            pause_button, language_entry, start_year_entry, end_year_entry, 
            keyword_entry, author_entry, file_type_entry)

def sort_tree(col, reverse):
    l = [(file_tree.set(k, col), k) for k in file_tree.get_children('')]
    l.sort(reverse=reverse)
    for index, (val, k) in enumerate(l):
        file_tree.move(k, '', index)
    file_tree.heading(col, command=lambda: sort_tree(col, not reverse))

# Create the GUI
(window, file_tree, download_dir_entry, 
 status_label, total_size_label, progress_bar, search_button, 
 pause_button, language_entry, start_year_entry, end_year_entry, 
 keyword_entry, author_entry, file_type_entry) = create_gui()

# Load search history
load_search_history()

# Start the main loop
window.protocol("WM_DELETE_WINDOW", lambda: [save_preferences(), window.destroy()])
window.mainloop()
