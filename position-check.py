import os
import pickle
import tkinter as tk
from tkinter import filedialog, messagebox
import webbrowser
import requests
from bs4 import BeautifulSoup
import pandas as pd
import time


def get_script_dir():
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return os.getcwd()


def load_settings():
    settings_file = os.path.join(get_script_dir(), 'settings.pkl')
    if os.path.exists(settings_file):
        with open(settings_file, 'rb') as f:
            return pickle.load(f)
    else:
        return {
            'user_agent': '',
            'language': 'da',  # Default: Danish
            'country': 'dk',   # Default: Denmark
            'queries': [],
            'target_site': '',  # Single site for search
            'proxies': []
        }


def save_settings(settings):
    settings_file = os.path.join(get_script_dir(), 'settings.pkl')
    with open(settings_file, 'wb') as f:
        pickle.dump(settings, f)


def search_google(settings):
    headers = {'User-Agent': settings['user_agent']}
    queries = settings['queries'][:100]
    target_site = settings['target_site']
    proxies = settings['proxies'] if settings['proxies'] else None
    current_proxy_index = 0

    results = []

    for query in queries:
        found = False
        while True:
            try:
                proxy = None
                if proxies:
                    proxy_entry = proxies[current_proxy_index]
                    proxy_url = proxy_entry['ip']
                    if proxy_entry['login'] and proxy_entry['password']:
                        proxy = {
                            'http': f"http://{proxy_entry['login']}:{proxy_entry['password']}@{proxy_url}",
                            'https': f"https://{proxy_entry['login']}:{proxy_entry['password']}@{proxy_url}"
                        }
                    else:
                        proxy = {
                            'http': f"http://{proxy_url}",
                            'https': f"https://{proxy_url}"
                        }

                url = f"https://www.google.com/search?q={query}&num=100&hl={settings['language']}&gl={settings['country']}"
                response = requests.get(url, headers=headers, proxies=proxy, timeout=5)

                # Обработка HTTP ошибок
                if response.status_code != 200:
                    print(f"HTTP Error: {response.status_code} for query '{query}'")
                    with open("debug.html", "w", encoding="utf-8") as file:
                        file.write(response.text)
                    raise Exception("Parsing error")

                soup = BeautifulSoup(response.text, 'html.parser')
                search_results = soup.find_all('div', class_='tF2Cxc')

                for index, result in enumerate(search_results, start=1):
                    link = result.find('a', href=True)
                    if link and target_site in link['href']:
                        found_url = link['href']
                        results.append({
                            'Query': query,
                            'Found URL': found_url,
                            'Position': index
                        })
                        print(f"Query: {query} | Found URL: {found_url} | Position: {index}")
                        found = True
                        break

                if not found:
                    results.append({
                        'Query': query,
                        'Found URL': 'Not found',
                        'Position': 'N/A'
                    })
                    print(f"Query: {query} | Found URL: Not found")
                break
            except Exception as e:
                if proxies:
                    current_proxy_index += 1
                    if current_proxy_index >= len(proxies):
                        messagebox.showerror("Error", f"Failed to parse query '{query}' with all proxies.")
                        return
                    continue
                else:
                    print(f"Error: {str(e)}")
                    messagebox.showerror("Error", f"Failed to parse query '{query}' without proxies.")
                    return

        # Добавляем задержку между запросами
        time.sleep(3)

    df = pd.DataFrame(results)
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")],
                                             initialdir=get_script_dir())
    if file_path:
        df.to_excel(file_path, index=False, sheet_name='Results')
        messagebox.showinfo("Information", "Results have been saved successfully.")


def create_gui(settings):
    window = tk.Tk()
    window.title("Search Engine Parser Settings")

    tk.Label(window, text="User Agent:").grid(row=0, column=0)
    user_agent_entry = tk.Entry(window, width=50)
    user_agent_entry.insert(0, settings['user_agent'])
    user_agent_entry.grid(row=0, column=1)

    link_label = tk.Label(window, text="Find your user agent at https://www.whatsmyua.info", fg="blue", cursor="hand2")
    link_label.grid(row=1, column=1)
    link_label.bind("<Button-1>", lambda e: webbrowser.open("https://www.whatsmyua.info"))

    tk.Label(window, text="Language (hl):").grid(row=2, column=0)
    language_entry = tk.Entry(window)
    language_entry.insert(0, settings['language'])
    language_entry.grid(row=2, column=1)

    tk.Label(window, text="Country (gl):").grid(row=3, column=0)
    country_entry = tk.Entry(window)
    country_entry.insert(0, settings['country'])
    country_entry.grid(row=3, column=1)

    tk.Label(window, text="Queries (up to 100, one per line):").grid(row=4, column=0)
    queries_text = tk.Text(window, height=10, width=50)
    queries_text.insert('1.0', '\n'.join(settings['queries']))
    queries_text.grid(row=4, column=1, columnspan=3)

    tk.Label(window, text="Target Site:").grid(row=5, column=0)
    target_site_entry = tk.Entry(window, width=50)
    target_site_entry.insert(0, settings['target_site'])
    target_site_entry.grid(row=5, column=1)

    tk.Label(window, text="Proxies (IP, Login, Password):").grid(row=6, column=0)
    proxy_frame = tk.Frame(window)
    proxy_frame.grid(row=6, column=1, columnspan=3)

    tk.Label(proxy_frame, text="IP Address", width=20, anchor="w").grid(row=0, column=0)
    tk.Label(proxy_frame, text="Login", width=15, anchor="w").grid(row=0, column=1)
    tk.Label(proxy_frame, text="Password", width=15, anchor="w").grid(row=0, column=2)

    proxy_entries = []
    for i in range(1, 6):  # Allow up to 5 proxies
        ip_entry = tk.Entry(proxy_frame, width=20)
        ip_entry.grid(row=i, column=0)

        login_entry = tk.Entry(proxy_frame, width=15)
        login_entry.grid(row=i, column=1)

        password_entry = tk.Entry(proxy_frame, width=15, show="*")
        password_entry.grid(row=i, column=2)

        proxy_entries.append({'ip': ip_entry, 'login': login_entry, 'password': password_entry})

    def on_save():
        settings['user_agent'] = user_agent_entry.get()
        settings['language'] = language_entry.get()
        settings['country'] = country_entry.get()
        settings['queries'] = queries_text.get('1.0', tk.END).splitlines()[:100]
        settings['target_site'] = target_site_entry.get()
        settings['proxies'] = [
            {
                'ip': entry['ip'].get(),
                'login': entry['login'].get(),
                'password': entry['password'].get()
            }
            for entry in proxy_entries if entry['ip'].get()
        ]
        save_settings(settings)
        search_google(settings)

    tk.Button(window, text="Save Settings and Run Search", command=on_save).grid(row=8, columnspan=2)

    window.mainloop()


if __name__ == "__main__":
    settings = load_settings()
    create_gui(settings)
