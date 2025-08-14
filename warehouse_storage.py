import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
import json
import os
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import pickle

class WarehouseStorage:
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    
    def __init__(self, parent=None):
        self.parent = parent
        self.service = None
        self.spreadsheet_id = None
        self.sheet_name = "Склад"
        self.storage_data = pd.DataFrame(columns=['Артикул', 'Количество', 'Ячейка'])
        self.storage_data.set_index('Артикул', inplace=True)
        
        # Файлы для сохранения настроек
        self.config_file = os.path.expanduser('~/.warehouse_storage_config.json')
        self.creds_file = r'E:\warehouse_storage\credentials.json'
        self.token_file = os.path.expanduser('~/.warehouse_storage_token.pickle')
        
        self.enabled = False
        self.load_config()
        
    def load_config(self):
        """Загрузка сохраненной конфигурации"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.spreadsheet_id = config.get('spreadsheet_id')
                    self.sheet_name = config.get('sheet_name', 'Склад')
                    self.enabled = config.get('enabled', False)
        except Exception as e:
            print(f"Ошибка загрузки конфигурации: {e}")
    
    def save_config(self):
        """Сохранение конфигурации"""
        try:
            config = {
                'spreadsheet_id': self.spreadsheet_id,
                'sheet_name': self.sheet_name,
                'enabled': self.enabled
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Ошибка сохранения конфигурации: {e}")
    
    def authenticate_google(self):
        """Аутентификация с Google Sheets API"""
        creds = None
        
        # Загрузка существующего токена
        if os.path.exists(self.token_file):
            try:
                with open(self.token_file, 'rb') as token:
                    creds = pickle.load(token)
            except:
                pass
        
        # Если нет валидных учетных данных, запрашиваем авторизацию
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                except:
                    creds = None
            
            if not creds:
                if not os.path.exists(self.creds_file):
                    messagebox.showerror(
                        "Ошибка", 
                        f"Не найден файл credentials.json.\n"
                        f"Поместите файл в: {self.creds_file}\n"
                        f"Получить можно в Google Cloud Console."
                    )
                    return False
                
                try:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        self.creds_file, self.SCOPES)
                    creds = flow.run_local_server(port=0)
                except Exception as e:
                    messagebox.showerror("Ошибка авторизации", str(e))
                    return False
            
            # Сохранение токена для следующего использования
            with open(self.token_file, 'wb') as token:
                pickle.dump(creds, token)
        
        try:
            self.service = build('sheets', 'v4', credentials=creds)
            return True
        except Exception as e:
            messagebox.showerror("Ошибка подключения", str(e))
            return False
    
    def create_spreadsheet_structure(self):
        """Создание структуры таблицы на Google Sheets"""
        if not self.service or not self.spreadsheet_id:
            return False
        
        try:
            # Получаем информацию о таблице
            sheet_metadata = self.service.spreadsheets().get(
                spreadsheetId=self.spreadsheet_id
            ).execute()
            
            # Проверяем, существует ли лист "Склад"
            sheet_exists = False
            for sheet in sheet_metadata.get('sheets', []):
                if sheet['properties']['title'] == self.sheet_name:
                    sheet_exists = True
                    break
            
            # Если лист не существует, создаем его
            if not sheet_exists:
                requests = [{
                    'addSheet': {
                        'properties': {
                            'title': self.sheet_name
                        }
                    }
                }]
                
                self.service.spreadsheets().batchUpdate(
                    spreadsheetId=self.spreadsheet_id,
                    body={'requests': requests}
                ).execute()
            
            # Устанавливаем заголовки
            headers = [['Артикул', 'Количество', 'Ячейка']]
            range_name = f'{self.sheet_name}!A1:C1'
            
            self.service.spreadsheets().values().update(
                spreadsheetId=self.spreadsheet_id,
                range=range_name,
                valueInputOption='RAW',
                body={'values': headers}
            ).execute()
            
            return True
            
        except Exception as e:
            messagebox.showerror("Ошибка создания структуры", str(e))
            return False
    
    def load_storage_data(self):
        """Загрузка данных склада из Google Sheets"""
        if not self.service or not self.spreadsheet_id:
            return False
        
        try:
            range_name = f'{self.sheet_name}!A:C'
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=range_name
            ).execute()
            
            values = result.get('values', [])
            if not values or len(values) < 2:
                # Пустая таблица или только заголовки
                self.storage_data = pd.DataFrame(columns=['Количество', 'Ячейка'])
                return True
            
            # Пропускаем заголовок
            data_rows = values[1:]
            
            # Обрабатываем данные
            processed_data = []
            for row in data_rows:
                if len(row) >= 2 and row[0].strip():  # Минимум артикул и количество
                    article = row[0].strip()
                    try:
                        quantity = int(row[1]) if row[1].strip() else 0
                    except:
                        quantity = 0
                    cell = row[2].strip() if len(row) > 2 else ""
                    processed_data.append({
                        'Артикул': article,
                        'Количество': quantity,
                        'Ячейка': cell
                    })
            
            if processed_data:
                self.storage_data = pd.DataFrame(processed_data)
                self.storage_data.set_index('Артикул', inplace=True)
            else:
                self.storage_data = pd.DataFrame(columns=['Количество', 'Ячейка'])
            
            return True
            
        except Exception as e:
            messagebox.showerror("Ошибка загрузки данных", str(e))
            return False
    
    def save_storage_data(self):
        """Сохранение данных склада в Google Sheets"""
        if not self.service or not self.spreadsheet_id:
            return False
        
        try:
            # Подготавливаем данные для записи
            values = [['Артикул', 'Количество', 'Ячейка']]
            
            for article, row in self.storage_data.iterrows():
                values.append([
                    str(article),
                    int(row['Количество']),
                    str(row['Ячейка'])
                ])
            
            # Очищаем существующие данные
            range_name = f'{self.sheet_name}!A:C'
            self.service.spreadsheets().values().clear(
                spreadsheetId=self.spreadsheet_id,
                range=range_name
            ).execute()
            
            # Записываем новые данные
            self.service.spreadsheets().values().update(
                spreadsheetId=self.spreadsheet_id,
                range=range_name,
                valueInputOption='RAW',
                body={'values': values}
            ).execute()
            
            return True
            
        except Exception as e:
            messagebox.showerror("Ошибка сохранения данных", str(e))
            return False
    
    def get_article_info(self, article):
        """Получение информации о товаре на складе"""
        if not self.enabled or article not in self.storage_data.index:
            return None, ""
        
        row = self.storage_data.loc[article]
        return row['Количество'], row['Ячейка']
    
    def update_article_quantity(self, article, quantity_change, cell=""):
        """Обновление количества товара на складе"""
        if not self.enabled:
            return True
        
        if article not in self.storage_data.index:
            # Добавляем новый товар
            self.storage_data.loc[article] = {'Количество': 0, 'Ячейка': cell}
        
        # Обновляем количество
        current_qty = self.storage_data.loc[article, 'Количество']
        new_qty = max(0, current_qty + quantity_change)  # Не допускаем отрицательные значения
        
        self.storage_data.loc[article, 'Количество'] = new_qty
        
        # Обновляем ячейку, если указана
        if cell:
            self.storage_data.loc[article, 'Ячейка'] = cell
        
        return True
    
    def show_storage_window(self):
        """Отображение окна управления складом"""
        storage_window = tk.Toplevel(self.parent)
        storage_window.title("Управление складом")
        storage_window.geometry("800x600")
        
        # Настройки подключения
        settings_frame = tk.LabelFrame(storage_window, text="Настройки Google Sheets")
        settings_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(settings_frame, text="ID таблицы:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        spreadsheet_entry = tk.Entry(settings_frame, width=50)
        spreadsheet_entry.grid(row=0, column=1, padx=5, pady=2)
        if self.spreadsheet_id:
            spreadsheet_entry.insert(0, self.spreadsheet_id)
        
        tk.Label(settings_frame, text="Название листа:").grid(row=1, column=0, sticky='w', padx=5, pady=2)
        sheet_entry = tk.Entry(settings_frame, width=30)
        sheet_entry.grid(row=1, column=1, sticky='w', padx=5, pady=2)
        sheet_entry.insert(0, self.sheet_name)
        
        # Кнопки управления
        btn_frame = tk.Frame(settings_frame)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        def connect_sheets():
            self.spreadsheet_id = spreadsheet_entry.get().strip()
            self.sheet_name = sheet_entry.get().strip() or "Склад"
            
            if not self.spreadsheet_id:
                messagebox.showerror("Ошибка", "Укажите ID таблицы")
                return
            
            if self.authenticate_google():
                if self.create_spreadsheet_structure():
                    if self.load_storage_data():
                        self.enabled = True
                        self.save_config()
                        messagebox.showinfo("Успех", "Подключение к Google Sheets установлено")
                        refresh_tree()
                    else:
                        messagebox.showerror("Ошибка", "Не удалось загрузить данные")
        
        def disconnect_sheets():
            self.enabled = False
            self.save_config()
            messagebox.showinfo("Информация", "Подключение к складу отключено")
            refresh_tree()
        
        tk.Button(btn_frame, text="Подключиться", command=connect_sheets).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Отключиться", command=disconnect_sheets).pack(side=tk.LEFT, padx=5)
        
        # Статус подключения
        status_label = tk.Label(settings_frame, text=f"Статус: {'Подключен' if self.enabled else 'Не подключен'}")
        status_label.grid(row=3, column=0, columnspan=2, pady=5)
        
        # Таблица данных склада
        data_frame = tk.LabelFrame(storage_window, text="Данные склада")
        data_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Поле для сканера
        scan_frame = tk.Frame(data_frame)
        scan_frame.pack(fill=tk.X, padx=5, pady=5)
        
        tk.Label(scan_frame, text="Артикул:").pack(side=tk.LEFT)
        article_entry = tk.Entry(scan_frame, width=20)
        article_entry.pack(side=tk.LEFT, padx=5)
        
        tk.Label(scan_frame, text="Количество:").pack(side=tk.LEFT, padx=(10,0))
        qty_entry = tk.Entry(scan_frame, width=10)
        qty_entry.pack(side=tk.LEFT, padx=5)
        qty_entry.insert(0, "1")
        
        tk.Label(scan_frame, text="Ячейка:").pack(side=tk.LEFT, padx=(10,0))
        cell_entry = tk.Entry(scan_frame, width=15)
        cell_entry.pack(side=tk.LEFT, padx=5)
        
        # Дерево для отображения данных
        tree = ttk.Treeview(data_frame, columns=("quantity", "cell"), show='tree headings')
        tree.heading('#0', text='Артикул')
        tree.heading('quantity', text='Количество')
        tree.heading('cell', text='Ячейка')
        tree.column('#0', width=200)
        tree.column('quantity', width=100)
        tree.column('cell', width=150)
        tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        def refresh_tree():
            for item in tree.get_children():
                tree.delete(item)
            
            if self.enabled and not self.storage_data.empty:
                for article, row in self.storage_data.iterrows():
                    tree.insert('', tk.END, text=article, 
                               values=(row['Количество'], row['Ячейка']))
        
        def add_item():
            article = article_entry.get().strip()
            if not article:
                messagebox.showwarning("Внимание", "Укажите артикул")
                return
            
            try:
                qty = int(qty_entry.get().strip())
            except:
                qty = 1
            
            cell = cell_entry.get().strip()
            
            if self.enabled:
                self.update_article_quantity(article, qty, cell)
                self.save_storage_data()
            
            article_entry.delete(0, tk.END)
            qty_entry.delete(0, tk.END)
            qty_entry.insert(0, "1")
            cell_entry.delete(0, tk.END)
            refresh_tree()
            article_entry.focus_set()
        
        def remove_item():
            selection = tree.selection()
            if not selection:
                messagebox.showwarning("Внимание", "Выберите товар для удаления")
                return
            
            article = tree.item(selection[0])['text']
            if messagebox.askyesno("Подтверждение", f"Удалить {article} со склада?"):
                if self.enabled and article in self.storage_data.index:
                    self.storage_data.drop(article, inplace=True)
                    self.save_storage_data()
                refresh_tree()
        
        # Кнопки управления данными
        data_btn_frame = tk.Frame(data_frame)
        data_btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        tk.Button(data_btn_frame, text="Добавить", command=add_item).pack(side=tk.LEFT, padx=5)
        tk.Button(data_btn_frame, text="Удалить", command=remove_item).pack(side=tk.LEFT, padx=5)
        tk.Button(data_btn_frame, text="Обновить", command=refresh_tree).pack(side=tk.LEFT, padx=5)
        
        # Привязка Enter к добавлению товара
        def on_enter(event):
            add_item()
        
        article_entry.bind('<Return>', on_enter)
        qty_entry.bind('<Return>', on_enter)
        cell_entry.bind('<Return>', on_enter)
        
        # Фокус на поле артикула
        article_entry.focus_set()
        
        # Инициализация таблицы
        refresh_tree()