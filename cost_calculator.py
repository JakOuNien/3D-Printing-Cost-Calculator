import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import json
import os
import sys
from collections import OrderedDict
import webbrowser

try:
    from PIL import Image, ImageTk
    from reportlab.pdfgen import canvas as pdfcanvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfbase import pdfmetrics
    LIBRARIES_INSTALLED = True
except ImportError:
    LIBRARIES_INSTALLED = False

# --- Language Manager Class ---
class LanguageManager:
    def __init__(self, language_file='languages.json'):
        self.languages = {}
        self.current_language = 'en'
        if os.path.exists(language_file):
            with open(language_file, 'r', encoding='utf-8-sig') as f:
                self.languages = json.load(f)

    def set_language(self, lang_code):
        if lang_code in self.languages:
            self.current_language = lang_code

    def get(self, key, lang_code=None):
        lang_to_use = lang_code if lang_code else self.current_language
        return self.languages.get(lang_to_use, {}).get(key, key)

    def get_available_languages(self):
        return list(self.languages.keys())

# --- Global Language Manager Instance ---
lang_manager = LanguageManager()

class Tooltip:
    def __init__(self, widget, text, bg='#333333', fg='white'):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.bg = bg
        self.fg = fg
        self.id = None
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
        self.widget.configure(cursor="question_arrow")

    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hide_tooltip()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(400, self.show_tooltip)

    def unschedule(self):
        id = self.id
        self.id = None
        if id:
            self.widget.after_cancel(id)

    def show_tooltip(self):
        if self.tooltip_window or not self.text: return
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.attributes("-alpha", 0.0)
        label = ttk.Label(self.tooltip_window, text=self.text, justify=tk.LEFT,
                          background=self.bg, foreground=self.fg, relief=tk.SOLID, borderwidth=1,
                          font=("Segoe UI", 9), wraplength=350, padding=(10, 8))
        label.pack(ipadx=1)
        self.tooltip_window.update_idletasks() 
        main_app_window = self.widget.winfo_toplevel()
        screen_width = main_app_window.winfo_width()
        screen_height = main_app_window.winfo_height()
        main_x = main_app_window.winfo_rootx()
        main_y = main_app_window.winfo_rooty()
        tip_width = self.tooltip_window.winfo_width()
        tip_height = self.tooltip_window.winfo_height()
        if x + tip_width > main_x + screen_width:
            x = main_x + screen_width - tip_width - 10
        if y + tip_height > main_y + screen_height:
            y = main_y + screen_height - tip_height - 10
        self.tooltip_window.wm_geometry(f"+{int(x)}+{int(y)}")
        self.fade_in()

    def fade_in(self, alpha=0):
        if self.tooltip_window:
            alpha = min(alpha + 0.08, 0.95)
            self.tooltip_window.attributes("-alpha", alpha)
            if alpha < 0.95:
                self.tooltip_window.after(15, lambda: self.fade_in(alpha))

    def hide_tooltip(self):
        if self.tooltip_window:
            self.tooltip_window.destroy()
        self.tooltip_window = None

class ScrollableFrame(ttk.Frame):
    def __init__(self, container, bg_color, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        self.canvas = tk.Canvas(self, highlightthickness=0, background=bg_color)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.scrollable_frame.bind("<Configure>", self.on_frame_configure)
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)

    def on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def _bind_mousewheel(self, event):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _unbind_mousewheel(self, event):
        self.canvas.unbind_all("<MouseWheel>")

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

class CalculationTab(ttk.Frame):
    def __init__(self, parent_notebook, main_app, **kwargs):
        super().__init__(parent_notebook, **kwargs)
        self.notebook = parent_notebook
        self.main_app = main_app
        self.variables = {}
        self.final_price = 0.0
        self.create_widgets()
        self.after(100, self.calculate_cost)

    def create_widgets(self):
        self.columnconfigure(0, weight=3, minsize=420)
        self.columnconfigure(1, weight=2, minsize=280)
        self.rowconfigure(0, weight=1)
        bg_color = self.main_app.active_theme['BG_COLOR']
        scrollable_input_area = ScrollableFrame(self, bg_color, style="Custom.TFrame")
        scrollable_input_area.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        input_frame = scrollable_input_area.scrollable_frame
        input_frame.configure(padding="10", style='TFrame')
        input_frame.columnconfigure(0, weight=1)
        result_frame = ttk.Frame(self, padding="20", style='TFrame')
        result_frame.grid(row=0, column=1, sticky="nsew")
        self.main_app.create_section(input_frame, "section_printing_params", "plastic", self)
        self.main_app.create_section(input_frame, "section_post_processing", "postprocess", self)
        self.main_app.create_section(input_frame, "section_packaging", "packaging", self)
        self.main_app.create_section(input_frame, "section_margin", "margin", self)
        self.main_app.create_section(input_frame, "section_adjustments", "adjustments", self)
        self.create_result_section(result_frame)

    def calculate_cost(self, *args):
        try:
            get_val = lambda name: self.variables.get(name, tk.DoubleVar(value=0)).get()
            plastic_cost = (get_val('total_weight_g') / 1000) * get_val('plastic_cost_kg')
            electricity_cost = (get_val('printer_power_w') / 1000) * get_val('print_time_h') * get_val('electricity_cost_kwh')
            depreciation_cost = get_val('print_time_h') * get_val('depreciation_h')
            print_subtotal = plastic_cost + electricity_cost + depreciation_cost
            failure_rate = get_val('failure_rate')
            print_total = print_subtotal / (1 - (failure_rate / 100)) if failure_rate < 100 else print_subtotal * 100
            post_process_labor_cost = ((get_val('support_removal_min') + get_val('sanding_min')) / 60) * get_val('labor_cost_h')
            post_process_total = post_process_labor_cost + get_val('consumables_cost')
            packaging_labor_cost = (get_val('packing_time_min') / 60) * get_val('labor_cost_h')
            packaging_total = get_val('packaging_cost_custom') + packaging_labor_cost
            cost_price = print_total + post_process_total + packaging_total
            profit_from_margin = cost_price * (get_val('margin_percentage') / 100)
            total_profit = profit_from_margin + get_val('fixed_profit')
            price_before_adjust = cost_price + total_profit
            discount_percentage = get_val('discount_percentage')
            additional_markup_percentage = get_val('additional_markup_percentage')
            if discount_percentage > 100: discount_percentage = 100
            final_price = price_before_adjust * (1 - (discount_percentage / 100))
            final_price = final_price * (1 + (additional_markup_percentage / 100))
            self.final_price = final_price
            
            set_val = lambda name, value: self.variables[name].set(f"{value:.2f} {self.main_app.currency}")
            set_val('result_plastic_cost', plastic_cost)
            set_val('result_electricity_cost', electricity_cost)
            set_val('result_depreciation', depreciation_cost)
            set_val('result_print_total', print_total)
            set_val('result_postprocessing_total', post_process_total)
            set_val('result_packaging_total', packaging_total)
            set_val('result_cost_price', cost_price)
            set_val('result_profit', total_profit)
            set_val('result_price_before_adjust', price_before_adjust)
            set_val('result_final_price', final_price)
            self.main_app.update_project_summary()
        except (tk.TclError, ValueError, KeyError, AttributeError): pass

    def create_result_section(self, parent):
        parent.columnconfigure(1, weight=1)
        self.result_labels = {}
        self.result_labels['header'] = ttk.Label(parent, text=lang_manager.get("result_header"), style='ResultHeader.TLabel')
        self.result_labels['header'].grid(row=0, column=0, columnspan=2, pady=(0, 15), sticky='w')
        r = 1
        self.result_labels['printing_subheader'] = ttk.Label(parent, text=lang_manager.get("result_printing_subheader"), style='ResultSubHeader.TLabel')
        self.result_labels['printing_subheader'].grid(row=r, column=0, columnspan=2, sticky='w', pady=(5,2)); r+=1
        r = self.main_app._create_result_row(parent, "result_plastic_cost", r, self);
        r = self.main_app._create_result_row(parent, "result_electricity_cost", r, self);
        r = self.main_app._create_result_row(parent, "result_depreciation", r, self);
        r = self.main_app._create_result_row(parent, "result_print_total", r, self, value_style='Header.TLabel');
        self.result_labels['postprocessing_subheader'] = ttk.Label(parent, text=lang_manager.get("result_postprocessing_subheader"), style='ResultSubHeader.TLabel')
        self.result_labels['postprocessing_subheader'].grid(row=r, column=0, columnspan=2, sticky='w', pady=(10,2)); r+=1
        r = self.main_app._create_result_row(parent, "result_postprocessing_total", r, self, value_style='Header.TLabel');
        self.result_labels['packaging_subheader'] = ttk.Label(parent, text=lang_manager.get("result_packaging_subheader"), style='ResultSubHeader.TLabel')
        self.result_labels['packaging_subheader'].grid(row=r, column=0, columnspan=2, sticky='w', pady=(10,2)); r+=1
        r = self.main_app._create_result_row(parent, "result_packaging_total", r, self, value_style='Header.TLabel');
        ttk.Separator(parent, orient='horizontal').grid(row=r, column=0, columnspan=2, sticky='ew', pady=15); r+=1
        r = self.main_app._create_result_row(parent, "result_cost_price", r, self, key_style='Header.TLabel');
        r = self.main_app._create_result_row(parent, "result_profit", r, self, key_style='Header.TLabel');
        r = self.main_app._create_result_row(parent, "result_price_before_adjust", r, self, key_style='Header.TLabel');
        ttk.Separator(parent, orient='horizontal').grid(row=r, column=0, columnspan=2, sticky='ew', pady=10); r+=1
        self.result_labels['final_price_header'] = ttk.Label(parent, text=lang_manager.get("result_final_price_header"), style='Total.TLabel')
        self.result_labels['final_price_header'].grid(row=r, column=0, sticky='w', pady=(15,0)); r+=1
        final_price_var = tk.StringVar(value=f"0.00 {self.main_app.currency}")
        self.variables['result_final_price'] = final_price_var
        ttk.Label(parent, textvariable=final_price_var, style='Total.TLabel').grid(row=r-1, column=1, sticky='e', pady=(15,0))

    def get_data(self):
        data = {name: var.get() for name, var in self.variables.items() if isinstance(var, tk.DoubleVar)}
        data['tab_title'] = self.notebook.tab(self, "text")
        for key in self.main_app.preset_fields:
            combo_key = f'{key}_combo'
            if combo_key in self.variables:
                data[combo_key] = self.variables[combo_key].get()
        return data

    def set_data(self, data):
        for name, value in data.items():
            if name.endswith('_combo'):
                if name in self.variables:
                    if value in self.variables[name]['values']:
                        self.variables[name].set(value)
            elif name in self.variables:
                self.variables[name].set(value)
        self.calculate_cost()
    
    def update_language(self):
        for key, label in self.result_labels.items():
            if not isinstance(key, str): continue
            label.config(text=lang_manager.get(key))
        for section in self.main_app.sections.get(self, {}).values():
            section['header_label'].config(text=lang_manager.get(section['lang_key']))
            for var_name, row_widgets in section['entry_rows'].items():
                row_widgets['label'].config(text=lang_manager.get(self.main_app.field_lang_keys.get(var_name, var_name)))


class InvoiceSettingsWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.transient(parent)
        self.title(lang_manager.get("invoice_settings_title"))
        self.parent = parent
        self.main_app = parent.main_app
        self.geometry("550x350")
        self.resizable(False, False)
        self.configure(bg=self.main_app.active_theme['BG_COLOR'])
        
        frame = ttk.Frame(self, padding=20, style="Custom.TFrame")
        frame.pack(fill="both", expand=True)
        
        self.settings_vars = {
            'shop_name': tk.StringVar(value=self.main_app.invoice_settings.get('shop_name', 'Your Company Name')),
            'logo_path': tk.StringVar(value=self.main_app.invoice_settings.get('logo_path', '')),
            'footer_text': tk.StringVar(value=self.main_app.invoice_settings.get('footer_text', 'Thank you for your order!'))
        }
        
        ttk.Label(frame, text=lang_manager.get("invoice_shop_name"), style="Custom.TLabel").grid(row=0, column=0, sticky="w", pady=10)
        ttk.Entry(frame, textvariable=self.settings_vars['shop_name'], width=40).grid(row=0, column=1, sticky="ew", pady=10)
        
        ttk.Label(frame, text=lang_manager.get("invoice_logo_path"), style="Custom.TLabel").grid(row=1, column=0, sticky="w", pady=10)
        logo_frame = ttk.Frame(frame, style="Custom.TFrame")
        logo_frame.grid(row=1, column=1, sticky="ew", pady=10)
        ttk.Entry(logo_frame, textvariable=self.settings_vars['logo_path'], width=30).pack(side="left", expand=True, fill="x")
        ttk.Button(logo_frame, text="...", width=3, command=self.select_logo, style="Small.TButton").pack(side="left", padx=(5,0))
        
        ttk.Label(frame, text=lang_manager.get("invoice_footer_text"), style="Custom.TLabel").grid(row=2, column=0, sticky="w", pady=10)
        ttk.Entry(frame, textvariable=self.settings_vars['footer_text'], width=40).grid(row=2, column=1, sticky="ew", pady=10)
        
        btn_frame = ttk.Frame(frame, style="Custom.TFrame")
        btn_frame.grid(row=3, column=0, columnspan=2, pady=(30, 0))
        ttk.Button(btn_frame, text=lang_manager.get("invoice_save"), command=self.save_and_close).pack(side="left", padx=5)
        ttk.Button(btn_frame, text=lang_manager.get("invoice_cancel"), command=self.destroy).pack(side="left", padx=5)
        
        self.grab_set()
        self.wait_window()

    def select_logo(self):
        path = filedialog.askopenfilename(title="Select Logo File", filetypes=[("Images", "*.png *.jpg *.jpeg"), ("All files", "*.*")])
        if path:
            self.settings_vars['logo_path'].set(path)
    
    def save_and_close(self):
        for key, var in self.settings_vars.items():
            self.main_app.invoice_settings[key] = var.get()
        self.main_app.save_config()
        self.parent.draw_preview()
        self.destroy()

class ModelInvoiceSettingsWindow(tk.Toplevel):
    def __init__(self, parent, model_name, current_settings):
        super().__init__(parent)
        self.transient(parent)
        self.title(f"{lang_manager.get('invoice_model_settings_title')} '{model_name}'")
        self.parent = parent
        self.model_name = model_name
        self.settings = current_settings.copy()
        self.main_app = parent.main_app
        self.geometry("500x600")
        self.configure(bg=self.main_app.active_theme['BG_COLOR'])

        main_frame = ttk.Frame(self, padding=15, style="Custom.TFrame")
        main_frame.pack(fill="both", expand=True)
        
        scroll_frame = ScrollableFrame(main_frame, self.main_app.active_theme['BG_COLOR'])
        scroll_frame.pack(fill="both", expand=True, pady=(0, 15))
        inner_frame = scroll_frame.scrollable_frame
        inner_frame.configure(style="Custom.TFrame", padding=5)

        self.vars = {}
        
        for group_key, group_data in self.main_app.invoice_detail_groups.items():
            group_container = ttk.Frame(inner_frame, style="Custom.TFrame")
            group_container.pack(fill='x', pady=(0, 15))
            
            ttk.Label(group_container, text=lang_manager.get(group_data['lang_key']), 
                      style="ResultSubHeader.TLabel").pack(anchor='w', pady=(0, 5))
            
            content_frame = ttk.Frame(group_container, style="Custom.TFrame", relief="solid", borderwidth=1)
            content_frame.pack(fill='x', ipadx=5, ipady=5)

            for item_key, item_lang_key in group_data['items'].items():
                self.vars[item_key] = tk.BooleanVar(value=self.settings.get(item_key, True))
                cb = ttk.Checkbutton(content_frame, text=lang_manager.get(item_lang_key), 
                                variable=self.vars[item_key], style="Custom.TCheckbutton")
                cb.pack(anchor='w', padx=10, pady=2)

        bottom_btn_frame = ttk.Frame(main_frame, style="Custom.TFrame")
        bottom_btn_frame.pack(fill='x')
        ttk.Button(bottom_btn_frame, text=lang_manager.get("invoice_save"), command=self.save_and_close).pack(side="right", padx=5)
        ttk.Button(bottom_btn_frame, text=lang_manager.get("invoice_cancel"), command=self.destroy).pack(side="right")
        
        self.grab_set()
        self.wait_window()

    def save_and_close(self):
        for key, var in self.vars.items():
            self.settings[key] = var.get()
        self.parent.update_model_settings(self.model_name, self.settings)
        self.destroy()

class InvoicePreviewWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.main_app = parent
        self.title(lang_manager.get("invoice_preview_title"))
        self.geometry("1100x850")
        self.configure(background=self.main_app.active_theme['BG_COLOR'])

        top_frame = ttk.Frame(self, padding=15, style="Custom.TFrame")
        top_frame.pack(side="top", fill="x")

        # --- Панель инструментов ---
        control_frame = ttk.Frame(top_frame, style="Custom.TFrame")
        control_frame.pack(side="left", fill="x", expand=True)
        
        ttk.Button(control_frame, text=lang_manager.get("invoice_save_pdf"), command=self.generate_pdf).pack(side="left", padx=(0, 10))
        ttk.Button(control_frame, text=lang_manager.get("invoice_settings"), command=self.open_settings).pack(side="left", padx=5)
        
        ttk.Label(control_frame, text=lang_manager.get("invoice_preset_label"), style="Custom.TLabel").pack(side="left", padx=(20, 5))
        self.preset_var = tk.StringVar(value="Standard")
        self.presets = {
            lang_manager.get("invoice_preset_simple"): "simple",
            lang_manager.get("invoice_preset_standard"): "standard",
            lang_manager.get("invoice_preset_transparent"): "transparent",
            lang_manager.get("invoice_preset_detailed"): "detailed"
        }
        preset_values = list(self.presets.keys())
        preset_combo = ttk.Combobox(control_frame, textvariable=self.preset_var, values=preset_values, state="readonly", width=25)
        preset_combo.pack(side="left")
        default_preset = lang_manager.get("invoice_preset_standard")
        if default_preset in preset_values:
            preset_combo.set(default_preset)
        elif preset_values:
            preset_combo.current(0)
            
        preset_combo.bind("<<ComboboxSelected>>", self.apply_invoice_preset)

        ttk.Label(control_frame, text=lang_manager.get("invoice_language"), style="Custom.TLabel").pack(side="left", padx=(20, 5))
        self.invoice_lang = tk.StringVar(value=lang_manager.current_language)
        lang_combo = ttk.Combobox(control_frame, textvariable=self.invoice_lang, values=lang_manager.get_available_languages(), state="readonly", width=5)
        lang_combo.pack(side="left")
        lang_combo.bind("<<ComboboxSelected>>", self.on_lang_change)
        
        ttk.Button(control_frame, text=lang_manager.get("invoice_close"), command=self.destroy).pack(side="right", padx=5)

        # --- Основная область ---
        container = ttk.Frame(self, style="Custom.TFrame")
        container.pack(fill="both", expand=True, padx=15, pady=10)
        
        left_pane = ScrollableFrame(container, self.main_app.active_theme['BG_COLOR'])
        left_pane.pack(side="left", fill="y", padx=(0, 15), ipadx=0)
        left_pane_inner = left_pane.scrollable_frame
        left_pane_inner.configure(style="Custom.TFrame")
        
        ttk.Label(left_pane_inner, text=lang_manager.get("invoice_select_models"), style="Header.TLabel").pack(anchor="w", pady=(0,10))
        
        self.model_vars = {}
        self.model_detail_settings = {}
        all_options = {key for group in self.main_app.invoice_detail_groups.values() for key in group['items']}

        for item in self.main_app.summary_tree.get_children():
            model_name, _ = self.main_app.summary_tree.item(item, 'values')
            self.model_detail_settings[model_name] = {key: True for key in all_options}
            
            model_frame = ttk.Frame(left_pane_inner, style="Custom.TFrame")
            model_frame.pack(anchor='w', fill='x', expand=True, pady=2)
            
            var = tk.BooleanVar(value=True)
            cb = ttk.Checkbutton(model_frame, text=model_name, variable=var, command=self.draw_preview, style="Custom.TCheckbutton")
            cb.pack(side='left', expand=True, fill='x')
            self.model_vars[model_name] = var
            
            settings_btn = ttk.Button(model_frame, text="⚙", style="Small.TButton", width=3,
                                      command=lambda m=model_name: self.open_model_settings(m))
            settings_btn.pack(side='right')

        ttk.Label(left_pane_inner, text=lang_manager.get("invoice_notes_label"), style="Header.TLabel").pack(anchor="w", pady=(20, 5))
        self.notes_text = tk.Text(left_pane_inner, height=5, width=35, font=("Segoe UI", 9), wrap="word", relief="solid", borderwidth=1)
        self.notes_text.pack(fill="x", padx=2)
        
        placeholder = lang_manager.get("invoice_notes_placeholder")
        self.notes_text.insert("1.0", placeholder)
        self.notes_text.config(fg="grey")
        self.notes_text.bind("<FocusIn>", lambda e: self.on_notes_focus_in(e, placeholder))
        self.notes_text.bind("<FocusOut>", lambda e: self.on_notes_focus_out(e, placeholder))
        self.notes_text.bind("<KeyRelease>", lambda e: self.draw_preview())

        self.preview_canvas = tk.Canvas(container, bg="#525659", highlightthickness=0)
        self.preview_canvas.pack(side="left", fill="both", expand=True)
        self.preview_canvas.bind("<Configure>", lambda e: self.draw_preview())
        
        self.apply_invoice_preset(None)

    def on_notes_focus_in(self, event, placeholder):
        if self.notes_text.get("1.0", "end-1c") == placeholder:
            self.notes_text.delete("1.0", "end")
            self.notes_text.config(fg="black")

    def on_notes_focus_out(self, event, placeholder):
        if not self.notes_text.get("1.0", "end-1c").strip():
            self.notes_text.insert("1.0", placeholder)
            self.notes_text.config(fg="grey")

    def apply_invoice_preset(self, event):
        selection = self.preset_var.get()
        preset_key = self.presets.get(selection, "standard")
        
        setting_map = {
            "simple": {
                "group_printing": False, "group_postprocessing": False, "group_packaging": False,
                "show_weight": False, "show_time": False,
                "show_cost_price": False, "show_profit": False, "show_price_before_adjust": False
            },
            "standard": {
                "group_printing": False, "group_postprocessing": True, "group_packaging": True,
                "show_weight": False, "show_time": False,
                "show_plastic_cost": False, "show_electricity_cost": False, "show_depreciation": False,
                "show_cost_price": False, "show_profit": False, "show_price_before_adjust": False
            },
            "transparent": {
                "show_weight": True, "show_time": True,
                "show_plastic_cost": True, "show_electricity_cost": True, "show_depreciation": True,
                "show_cost_price": True,
                "show_profit": False,
                "show_price_before_adjust": False
            },
            "detailed": {}
        }

        new_conf = setting_map.get(preset_key, {})
        all_options = {key for group in self.main_app.invoice_detail_groups.values() for key in group['items']}
        
        for model_name in self.model_detail_settings:
            for key in all_options:
                if preset_key == "detailed":
                    self.model_detail_settings[model_name][key] = True
                elif preset_key == "simple":
                    self.model_detail_settings[model_name][key] = False
                else: 
                    if key in new_conf:
                        self.model_detail_settings[model_name][key] = new_conf[key]
                    else:
                        self.model_detail_settings[model_name][key] = True

        self.draw_preview()

    def open_model_settings(self, model_name):
        ModelInvoiceSettingsWindow(self, model_name, self.model_detail_settings[model_name])

    def update_model_settings(self, model_name, new_settings):
        self.model_detail_settings[model_name] = new_settings
        self.draw_preview()

    def on_lang_change(self, event=None):
        self.draw_preview()
        
    def open_settings(self):
        InvoiceSettingsWindow(self)

    def draw_preview(self, event=None):
        self.preview_canvas.delete("all")
        canvas_w = self.preview_canvas.winfo_width()
        canvas_h = self.preview_canvas.winfo_height()
        if canvas_w < 10 or canvas_h < 10: return
        
        a4_w, a4_h = A4
        scale = min((canvas_w - 40) / a4_w, (canvas_h - 40) / a4_h)
        
        page_w = a4_w * scale
        page_h = a4_h * scale
        x_off = (canvas_w - page_w) / 2
        y_off = 20
        
        self.preview_canvas.create_rectangle(x_off, y_off, x_off + page_w, y_off + page_h, fill="white", outline="#000000")
        self.preview_canvas.create_rectangle(x_off + 2, y_off + 2, x_off + page_w + 2, y_off + page_h + 2, fill="#000000", stipple="gray25", outline="")

        margin = (1.5 * cm) * scale
        current_y = y_off + margin
        
        font_title = ("Arial", int(14 * scale), "bold")
        font_normal = ("Arial", int(10 * scale))
        font_bold = ("Arial", int(10 * scale), "bold")
        font_small = ("Arial", int(8 * scale))
        
        self.preview_canvas.create_text(x_off + margin, current_y, text=self.main_app.invoice_settings.get('shop_name', ''), anchor="nw", font=font_title)
        current_y += 40 * scale
        self.preview_canvas.create_line(x_off + margin, current_y, x_off + page_w - margin, current_y, fill="#CCCCCC")
        current_y += 20 * scale
        
        lang_code = self.invoice_lang.get()
        selected_models = [name for name, var in self.model_vars.items() if var.get()]
        
        for model_name in selected_models:
            price_str = f"0.00 {self.main_app.currency}"
            for item in self.main_app.summary_tree.get_children():
                if self.main_app.summary_tree.item(item, 'values')[0] == model_name:
                    price_str = self.main_app.summary_tree.item(item, 'values')[1]
                    break
            
            self.preview_canvas.create_rectangle(x_off + margin, current_y, x_off + page_w - margin, current_y + (25*scale), fill="#333333", outline="")
            self.preview_canvas.create_text(x_off + margin + (5*scale), current_y + (12*scale), text=model_name, anchor="w", font=font_bold, fill="white")
            self.preview_canvas.create_text(x_off + page_w - margin - (5*scale), current_y + (12*scale), text=price_str, anchor="e", font=font_bold, fill="white")
            current_y += 35 * scale
            
            tab_widget = None
            for t_id in self.main_app.notebook.tabs():
                if self.main_app.notebook.tab(t_id, "text") == model_name:
                    tab_widget = self.main_app.nametowidget(t_id)
                    break
            
            settings = self.model_detail_settings.get(model_name, {})
            
            if tab_widget:
                for group_key, group_data in self.main_app.invoice_detail_groups.items():
                    has_visible_items = any(settings.get(k, True) for k in group_data['items'])
                    
                    show_group_header = True
                    if self.preset_var.get() == lang_manager.get("invoice_preset_simple"):
                        show_group_header = False

                    if has_visible_items or (group_data['total_key'] and show_group_header):
                         if group_data['total_key'] and show_group_header:
                             header_txt = lang_manager.get(group_data['lang_key'], lang_code)
                             total_val = tab_widget.variables.get(group_data['total_key'], tk.StringVar(value="")).get()
                             
                             self.preview_canvas.create_text(x_off + margin, current_y, text=header_txt, anchor="w", font=font_bold)
                             self.preview_canvas.create_text(x_off + page_w - margin, current_y, text=total_val, anchor="e", font=font_bold)
                             current_y += 15 * scale
                         
                         for item_key, item_lang_key in group_data['items'].items():
                             if settings.get(item_key, True):
                                 raw_val = tab_widget.variables[item_lang_key].get()
                                 
                                 display_val = str(raw_val)
                                 if item_key == 'show_weight':
                                     try: display_val = f"{float(raw_val):.0f} g"
                                     except: pass
                                 elif item_key == 'show_time':
                                     try: display_val = f"{float(raw_val):.1f} h"
                                     except: pass

                                 label = lang_manager.get(item_lang_key, lang_code)
                                 if item_key == 'show_weight': label = lang_manager.get('param_weight', lang_code)
                                 if item_key == 'show_time': label = lang_manager.get('param_time', lang_code)

                                 self.preview_canvas.create_text(x_off + margin + (10*scale), current_y, text=lbl, anchor="w", font=font_normal, fill="#555555")
                                 self.preview_canvas.create_text(x_off + page_w - margin, current_y, text=display_val, anchor="e", font=font_normal, fill="#555555")
                                 current_y += 15 * scale
                         
                         if has_visible_items:
                             current_y += 5 * scale

            current_y += 10 * scale

        footer_y = y_off + page_h - margin - (15 * scale)
        
        notes_content = self.notes_text.get("1.0", "end-1c")
        placeholder = lang_manager.get("invoice_notes_placeholder")
        if notes_content and notes_content != placeholder and notes_content.strip():
             self.preview_canvas.create_text(x_off + margin, current_y + (20*scale), text=lang_manager.get("invoice_notes_label", lang_code), font=font_bold, anchor="w")
             self.preview_canvas.create_text(x_off + margin, current_y + (35*scale), text=notes_content, font=font_small, anchor="nw", width=(page_w - 2*margin))

        total = self.main_app.summary_vars['total'].get()
        grand_total = self.main_app.summary_vars['grand_total'].get()
        
        ty = current_y + 80 * scale
        self.preview_canvas.create_line(x_off + page_w/2, ty, x_off + page_w - margin, ty, fill="#000000")
        ty += 15 * scale
        
        self.preview_canvas.create_text(x_off + page_w - margin - (80*scale), ty, text=lang_manager.get('pdf_total_amount', lang_code) + ":", anchor="e", font=font_normal)
        self.preview_canvas.create_text(x_off + page_w - margin, ty, text=total, anchor="e", font=font_normal)
        ty += 20 * scale
        
        self.preview_canvas.create_text(x_off + page_w - margin - (80*scale), ty, text=lang_manager.get('pdf_total_due', lang_code) + ":", anchor="e", font=font_title)
        self.preview_canvas.create_text(x_off + page_w - margin, ty, text=grand_total, anchor="e", font=font_title, fill=self.main_app.active_theme['ACCENT_COLOR'])

        self.preview_canvas.create_text(x_off + page_w/2, footer_y, text=self.main_app.invoice_settings.get('footer_text', ''), anchor="center", font=font_small, fill="#777777")

    def generate_pdf(self):
        selected_models = [name for name, var in self.model_vars.items() if var.get()]
        if not selected_models:
            messagebox.showwarning("No Models Selected", "Please select at least one model.")
            return

        filepath = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Documents", "*.pdf")], title="Save Invoice As...")
        if not filepath: return
        
        notes = self.notes_text.get("1.0", "end-1c")
        if notes == lang_manager.get("invoice_notes_placeholder"): notes = ""
        
        try:
            self.main_app.generate_invoice_pdf(filepath, selected_models, self.invoice_lang.get(), self.model_detail_settings, notes)
            messagebox.showinfo("Success", f"Invoice saved successfully!\n{filepath}")
            if messagebox.askyesno("Open File", "Open PDF now?"):
                webbrowser.open(filepath)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create PDF: {e}")

class CostCalculatorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.config_file = 'config.json'
        self.presets_file = 'presets.json'
        self.config_data = self.load_config()
        lang_manager.set_language(self.config_data.get('language', 'en'))
        
        self.currency = self.config_data.get("currency", "€")
        self.available_currencies = ["€", "$", "£", "₴", "¥", "zł"]

        if LIBRARIES_INSTALLED and os.path.exists('DejaVuSans.ttf') and os.path.exists('DejaVuSans-Bold.ttf'):
            pdfmetrics.registerFont(TTFont('DejaVu', 'DejaVuSans.ttf'))
            pdfmetrics.registerFont(TTFont('DejaVu-Bold', 'DejaVuSans-Bold.ttf'))
        else:
            messagebox.showwarning(
                "Font Missing",
                "For Cyrillic support, ensure DejaVuSans.ttf is in the folder."
            )

        self.title(lang_manager.get("app_title"))
        self.geometry("1400x900")
        self.state('zoomed')
        self.minsize(1024, 768)

        self.invoice_settings = self.config_data.get("invoice_settings", {})
        self.custom_themes = self.config_data.get("custom_themes", {})
        self.invoice_detail_presets = self.config_data.get("invoice_detail_presets", {})
        self.sections = {} 
        self.init_themes()
        for name, path in self.custom_themes.items():
            if os.path.exists(path):
                try:
                    with open(path, 'r', encoding='utf-8-sig') as f: self.themes[name] = json.load(f)
                except Exception: pass
        self.init_field_lang_keys()
        self.init_invoice_detail_groups()
        self.init_help_texts()
        self.init_preset_fields()
        self.presets = self.load_presets()
        self.active_theme_name = self.config_data.get("theme", "Спокойная Мята")
        self.active_theme = self.themes.get(self.active_theme_name, self.themes["Спокойная Мята"])
        self.setup_styles()
        self.create_main_layout()
        self.create_menu()
        self.add_new_tab()

    def get_default_themes(self):
        return ["Спокойная Мята", "Графит"]

    def init_themes(self):
        self.themes = {
            "Спокойная Мята": {"BG_COLOR": "#f0f5f5", "FRAME_COLOR": "#ffffff", "TEXT_ON_BG_COLOR": "#2a3f54", "TEXT_ON_FRAME_COLOR": "#2a3f54", "ACCENT_COLOR": "#2a7f9a", "ACCENT_HOVER_COLOR": "#1f6a83", "BORDER_COLOR": "#dbe3e3", "TAB_BG": "#d9e0e0", "TAB_FG": "#4b5563", "TAB_SELECTED_BG": "#ffffff", "TAB_SELECTED_FG": "#2a7f9a", "TREE_HEADING_BG": "#d9e0e0", "TREE_FIELD_BG": "#ffffff", "COMBO_LIST_BG": "#ffffff", "COMBO_LIST_FG": "#2a3f54", "COMBO_LIST_SELECT_BG": "#2a7f9a", "COMBO_LIST_SELECT_FG": "#ffffff", "BUTTON_TEXT_COLOR": "#ffffff", "SELECTED_TEXT_COLOR": "#ffffff", "COMBO_FIELD_FG": "#2a3f54"},
            "Графит": {"BG_COLOR": "#263238", "FRAME_COLOR": "#37474f", "TEXT_ON_BG_COLOR": "#eceff1", "TEXT_ON_FRAME_COLOR": "#eceff1", "ACCENT_COLOR": "#00bcd4", "ACCENT_HOVER_COLOR": "#0097a7", "BORDER_COLOR": "#546e7a", "TAB_BG": "#263238", "TAB_FG": "#b0bec5", "TAB_SELECTED_BG": "#37474f", "TAB_SELECTED_FG": "#00bcd4", "TREE_HEADING_BG": "#546e7a", "TREE_FIELD_BG": "#455a64", "COMBO_LIST_BG": "#455a64", "COMBO_LIST_FG": "#eceff1", "COMBO_LIST_SELECT_BG": "#00bcd4", "COMBO_LIST_SELECT_FG": "#263238", "BUTTON_TEXT_COLOR": "#263238", "SELECTED_TEXT_COLOR": "#263238", "COMBO_FIELD_FG": "#263238"}
        }
        self.tooltip_bg = '#37474f'
        self.tooltip_fg = '#eceff1'

    def setup_styles(self):
        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        c = self.active_theme
        self.configure(background=c['BG_COLOR'])
        self.option_add("*TCombobox*Listbox*Background", c['COMBO_LIST_BG']); self.option_add("*TCombobox*Listbox*Foreground", c['COMBO_LIST_FG']); self.option_add("*TCombobox*Listbox*selectBackground", c['COMBO_LIST_SELECT_BG']); self.option_add("*TCombobox*Listbox*selectForeground", c['COMBO_LIST_SELECT_FG'])
        self.style.configure('.', background=c['BG_COLOR'], foreground=c['TEXT_ON_BG_COLOR'], font=('Segoe UI', 10), bordercolor=c['BORDER_COLOR'])
        self.style.configure('TNotebook', background=c['BG_COLOR'], borderwidth=0)
        self.style.configure('TNotebook.Tab', background=c['TAB_BG'], foreground=c['TAB_FG'], padding=[12, 6], font=('Segoe UI', 10, 'bold'), borderwidth=0)
        self.style.map('TNotebook.Tab', background=[('selected', c['TAB_SELECTED_BG'])], foreground=[('selected', c['TAB_SELECTED_FG'])])
        self.style.configure('TLabel', background=c['FRAME_COLOR'], foreground=c['TEXT_ON_FRAME_COLOR'])
        self.style.configure("Custom.TLabel", background=c['BG_COLOR'], foreground=c['TEXT_ON_BG_COLOR'])
        self.style.configure("Custom.TCheckbutton", background=c['BG_COLOR'], foreground=c['TEXT_ON_BG_COLOR'])
        self.style.configure("Custom.TLabelframe", background=c['BG_COLOR'], relief="groove", borderwidth=1, bordercolor=c['BORDER_COLOR'])
        self.style.configure("Custom.TLabelframe.Label", background=c['BG_COLOR'], foreground=c['TEXT_ON_BG_COLOR'])
        self.style.map('Custom.TCheckbutton', indicatorbackground=[('selected', c['ACCENT_COLOR']), ('active', c['FRAME_COLOR'])], background=[('active', c['BG_COLOR'])])
        self.style.configure('Header.TLabel', font=('Segoe UI', 11, 'bold'), background=c['BG_COLOR'], foreground=c['ACCENT_COLOR'])
        self.style.configure('ResultHeader.TLabel', font=('Segoe UI', 13, 'bold'), background=c['FRAME_COLOR'], foreground=c['TEXT_ON_FRAME_COLOR'])
        self.style.configure('ResultSubHeader.TLabel', font=('Segoe UI', 11, 'bold'), background=c['FRAME_COLOR'], foreground=c['ACCENT_COLOR'])
        self.style.configure('Total.TLabel', font=('Segoe UI', 18, 'bold'), background=c['FRAME_COLOR'], foreground=c['ACCENT_COLOR'])
        self.style.configure('ResultKey.TLabel', background=c['FRAME_COLOR'], foreground=c['TEXT_ON_FRAME_COLOR'])
        self.style.configure('ResultValue.TLabel', background=c['FRAME_COLOR'], foreground=c['TEXT_ON_FRAME_COLOR'])
        self.style.configure('TButton', background=c['ACCENT_COLOR'], foreground=c['BUTTON_TEXT_COLOR'], font=('Segoe UI', 9, 'bold'), borderwidth=0, padding=6)
        self.style.map('TButton', background=[('active', c['ACCENT_HOVER_COLOR'])])
        self.style.configure('Small.TButton', font=('Segoe UI', 8), padding=4)
        self.style.configure('Help.TButton', font=('Segoe UI', 8, 'bold'), padding=(2,0), relief='flat', borderwidth=0, background=c['FRAME_COLOR'], foreground='#9ca3af')
        self.style.map('Help.TButton', foreground=[('active', c['ACCENT_COLOR'])])
        self.style.configure('TEntry', fieldbackground=c['FRAME_COLOR'], bordercolor=c['BORDER_COLOR'], foreground=c['TEXT_ON_FRAME_COLOR'], insertcolor=c['TEXT_ON_FRAME_COLOR'])
        self.style.configure('TCombobox', fieldbackground=c['FRAME_COLOR'], foreground=c['COMBO_FIELD_FG'], arrowcolor=c['TEXT_ON_FRAME_COLOR'], insertcolor=c['TEXT_ON_FRAME_COLOR'])
        self.style.configure('TFrame', background=c['FRAME_COLOR'])
        self.style.configure("Custom.TFrame", background=c['BG_COLOR'])
        self.style.configure("Treeview", rowheight=25, fieldbackground=c['TREE_FIELD_BG'], background=c['TREE_FIELD_BG'], foreground=c['TEXT_ON_FRAME_COLOR'])
        self.style.map("Treeview", background=[('selected', c['ACCENT_COLOR'])], foreground=[('selected', c['SELECTED_TEXT_COLOR'])])
        self.style.configure("Treeview.Heading", font=('Segoe UI', 10, 'bold'), background=c['TREE_HEADING_BG'], foreground=c['TEXT_ON_FRAME_COLOR'])
        self.style.configure("Vertical.TScrollbar", background=c['FRAME_COLOR'])

    def init_preset_fields(self):
        self.preset_fields = { "plastic": OrderedDict([('plastic_cost_kg', 25.0)]), "postprocess": OrderedDict([('labor_cost_h', 15.0), ('support_removal_min', 15.0), ('sanding_min', 0.0), ('consumables_cost', 0.0)]), "packaging": OrderedDict([('packaging_cost_custom', 2.0), ('packing_time_min', 5.0)]), "margin": OrderedDict([('margin_percentage', 100.0), ('fixed_profit', 0.0)]) }
    
    def init_field_lang_keys(self):
        self.field_lang_keys = { 'plastic_cost_kg': "result_plastic_cost", 'total_weight_g': "total_weight_g", 'support_percentage': "support_percentage", 'print_time_h': "print_time_h", 'electricity_cost_kwh': "result_electricity_cost", 'printer_power_w': "printer_power_w", 'depreciation_h': "result_depreciation", 'failure_rate': "failure_rate", 'labor_cost_h': "labor_cost_h", 'support_removal_min': "support_removal_min", 'sanding_min': "sanding_min", 'consumables_cost': "consumables_cost", 'packaging_cost_custom': "packaging_cost_custom", 'packing_time_min': "packing_time_min", 'margin_percentage': "margin_percentage", 'fixed_profit': "fixed_profit", 'discount_percentage': "discount_percentage", 'additional_markup_percentage': "additional_markup_percentage", 'overall_discount': "overall_discount", 'overall_markup': "overall_markup" }

    def init_invoice_detail_groups(self):
        self.invoice_detail_groups = OrderedDict([
            ("group_printing", {
                "lang_key": "group_printing",
                "total_key": "result_print_total",
                "items": OrderedDict([
                    ('show_weight', 'total_weight_g'),
                    ('show_time', 'print_time_h'),
                    ('show_plastic_cost', 'result_plastic_cost'),
                    ('show_electricity_cost', 'result_electricity_cost'),
                    ('show_depreciation', 'result_depreciation')
                ])
            }),
            ("group_postprocessing", {
                "lang_key": "group_postprocessing",
                "total_key": "result_postprocessing_total",
                "items": OrderedDict() 
            }),
            ("group_packaging", {
                "lang_key": "group_packaging",
                "total_key": "result_packaging_total",
                "items": OrderedDict()
            }),
            ("group_final", {
                "lang_key": "group_final",
                "total_key": None,
                "items": OrderedDict([
                    ('show_cost_price', 'result_cost_price'),
                    ('show_profit', 'result_profit'),
                    ('show_price_before_adjust', 'result_price_before_adjust')
                ])
            })
        ])

    def init_help_texts(self): 
        self.help_texts = {} 

    def get_default_presets(self): return { "plastic": { "PETG (Стандарт)": {"plastic_cost_kg": 25.0}}, "postprocess": { "Только поддержки": {"labor_cost_h": 15.0}}, "packaging": { "Пакет + пупырка": {"packaging_cost_custom": 1.0}}, "margin": { "Стандарт (x2)": {"margin_percentage": 100.0}}}
    
    def load_presets(self):
        defaults = self.get_default_presets()
        if os.path.exists(self.presets_file):
            try:
                with open(self.presets_file, 'r', encoding='utf-8-sig') as f:
                    loaded = json.load(f)
                    for cat in defaults:
                        if cat not in loaded: loaded[cat] = defaults[cat]
                    return loaded
            except (json.JSONDecodeError, IOError): return defaults
        return defaults

    def save_presets_to_file(self):
        try:
            with open(self.presets_file, 'w', encoding='utf-8') as f: json.dump(self.presets, f, indent=4, ensure_ascii=False)
        except IOError as e: messagebox.showerror("Error", f"Failed to save presets: {e}")

    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8-sig') as f: return json.load(f)
            except Exception: return {}
        return {}

    def save_config(self):
        try:
            self.config_data['invoice_settings'] = self.invoice_settings
            self.config_data['custom_themes'] = self.custom_themes
            self.config_data['language'] = lang_manager.current_language
            self.config_data['currency'] = self.currency
            self.config_data['invoice_detail_presets'] = self.invoice_detail_presets
            with open(self.config_file, 'w', encoding='utf-8') as f: json.dump(self.config_data, f, indent=4, ensure_ascii=False)
        except IOError: pass

    def change_theme(self, theme_name, custom_theme_data=None):
        if custom_theme_data: self.themes[theme_name] = custom_theme_data
        if theme_name in self.themes:
            self.active_theme_name = theme_name
            self.active_theme = self.themes[theme_name]
            self.setup_styles()
            self.config_data['theme'] = self.active_theme_name
            self.save_config()
            self.rebuild_view_menu()
            
    def change_currency(self, new_currency):
        self.currency = new_currency
        self.save_config()
        for tab_id in self.notebook.tabs():
            tab_widget = self.nametowidget(tab_id)
            tab_widget.calculate_cost()
        self.update_project_summary()

    def load_custom_theme(self, filepath=None, show_success_message=True):
        if not filepath: filepath = filedialog.askopenfilename(title="Select Theme File", filetypes=[("JSON Theme Files", "*.json")])
        if not filepath: return
        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f: custom_theme = json.load(f)
            if not all(key in custom_theme for key in self.themes["Графит"].keys()):
                messagebox.showerror("Theme Error", "Theme file has incorrect structure.")
                return
            theme_name = os.path.basename(filepath).split('.')[0]
            if theme_name in self.get_default_themes():
                messagebox.showerror("Error", f"Theme name '{theme_name}' is reserved.")
                return
            self.custom_themes[theme_name] = filepath
            self.change_theme(theme_name, custom_theme_data=custom_theme)
            if show_success_message: messagebox.showinfo("Success", f"Theme '{theme_name}' loaded successfully.")
        except Exception as e:
            if show_success_message: messagebox.showerror("Error", f"Failed to load theme: {e}")

    def create_menu(self):
        self.menu_bar = tk.Menu(self)
        self.config(menu=self.menu_bar)
        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label=lang_manager.get("menu_file"), menu=self.file_menu)
        self.view_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label=lang_manager.get("menu_view"), menu=self.view_menu)
        self.lang_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label=lang_manager.get("menu_language"), menu=self.lang_menu)
        self.currency_menu = tk.Menu(self.menu_bar, tearoff=0)
        
        cur_label = lang_manager.get("menu_currency")
        if cur_label == "menu_currency":
            cur_label = "Валюта" if lang_manager.current_language == 'ru' else "Currency"
            
        self.menu_bar.add_cascade(label=cur_label, menu=self.currency_menu)
        self.update_language_menu_texts()

    def set_language(self, lang_code):
        if lang_code == lang_manager.current_language:
            return
            
        lang_manager.set_language(lang_code)
        self.save_config()
        
        msg_title = "Перезапуск" if lang_code == 'ru' else "Restart Required"
        msg_body = "Приложение будет перезапущено для применения языка." if lang_code == 'ru' else "The application will now restart to apply the language change."
        messagebox.showinfo(msg_title, msg_body)
        
        save_title = "Сохранение" if lang_code == 'ru' else "Save Progress"
        save_body = "Вы хотите сохранить текущий проект перед перезапуском?" if lang_code == 'ru' else "Do you want to save your project before restarting?"
        
        save_choice = messagebox.askyesno(save_title, save_body)
        if save_choice:
            saved = self.save_project()
            if not saved:
                return
                
        self.restart_app()
        
    def restart_app(self):
        self.destroy()
        os.execl(sys.executable, sys.executable, *sys.argv)
    
    def update_language_menu_texts(self):
        self.file_menu.delete(0, tk.END)
        self.file_menu.add_command(label=lang_manager.get("menu_new_project"), command=self.new_project)
        self.file_menu.add_command(label=lang_manager.get("menu_open_project"), command=self.load_project)
        self.file_menu.add_command(label=lang_manager.get("menu_save_project"), command=self.save_project)
        self.file_menu.add_separator()
        self.file_menu.add_command(label=lang_manager.get("menu_exit"), command=self.quit)
        self.menu_bar.entryconfigure(1, label=lang_manager.get("menu_file"))
        self.menu_bar.entryconfigure(2, label=lang_manager.get("menu_view"))
        self.menu_bar.entryconfigure(3, label=lang_manager.get("menu_language"))
        
        cur_label = lang_manager.get("menu_currency")
        if cur_label == "menu_currency":
            cur_label = "Валюта" if lang_manager.current_language == 'ru' else "Currency"
        self.menu_bar.entryconfigure(4, label=cur_label)
        
        self.rebuild_view_menu()
        
        self.lang_menu.delete(0, tk.END)
        for lang_code in lang_manager.get_available_languages():
            self.lang_menu.add_command(label=lang_code.upper(), command=lambda lc=lang_code: self.set_language(lc))
            
        self.currency_menu.delete(0, tk.END)
        for cur in self.available_currencies:
            self.currency_menu.add_command(label=cur, command=lambda c=cur: self.change_currency(c))

    def rebuild_view_menu(self):
        self.view_menu.delete(0, tk.END)
        for theme_name in sorted(self.get_default_themes()):
            self.view_menu.add_command(label=theme_name, command=lambda t=theme_name: self.change_theme(t))
        if self.custom_themes:
            self.view_menu.add_separator()
            for theme_name in sorted(self.custom_themes.keys()):
                self.view_menu.add_command(label=f"{theme_name} (file)", command=lambda t=theme_name: self.change_theme(t))
        self.view_menu.add_separator()
        self.view_menu.add_command(label=lang_manager.get("menu_load_theme"), command=self.load_custom_theme)

    def create_main_layout(self):
        self.main_pane = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        self.main_pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        notebook_frame = ttk.Frame(self.main_pane, style="Custom.TFrame")
        notebook_frame.rowconfigure(1, weight=1)
        notebook_frame.columnconfigure(0, weight=1)
        tab_actions_frame = ttk.Frame(notebook_frame, style="Custom.TFrame")
        tab_actions_frame.grid(row=0, column=0, sticky="ew", pady=(0,5))
        self.tab_add_btn = ttk.Button(tab_actions_frame, text=lang_manager.get("tab_add"), command=self.add_new_tab, width=4)
        self.tab_add_btn.pack(side=tk.LEFT)
        self.tab_rename_btn = ttk.Button(tab_actions_frame, text=lang_manager.get("tab_rename"), command=self.rename_current_tab, style='Small.TButton')
        self.tab_rename_btn.pack(side=tk.LEFT, padx=5)
        self.tab_delete_btn = ttk.Button(tab_actions_frame, text=lang_manager.get("tab_delete"), command=self.delete_current_tab, style='Small.TButton')
        self.tab_delete_btn.pack(side=tk.LEFT)
        self.notebook = ttk.Notebook(notebook_frame)
        self.notebook.grid(row=1, column=0, sticky="nsew")
        self.main_pane.add(notebook_frame, weight=3)
        self.summary_frame = ttk.Frame(self.main_pane, style='TFrame', relief='solid', borderwidth=1, padding=20)
        self.create_project_summary_widgets(self.summary_frame)
        self.main_pane.add(self.summary_frame, weight=1)

    def create_project_summary_widgets(self, parent):
        self.summary_header_label = ttk.Label(parent, text=lang_manager.get("summary_header"), style='ResultHeader.TLabel')
        self.summary_header_label.pack(anchor='w', pady=(0, 15))
        cols = ('model', 'price'); self.summary_tree = ttk.Treeview(parent, columns=cols, show='headings', height=10)
        self.summary_tree.pack(fill=tk.BOTH, expand=True)
        self.update_summary_tree_headings()
        ttk.Separator(parent, orient='horizontal').pack(fill='x', pady=15)
        summary_total_frame = ttk.Frame(parent, style='TFrame'); summary_total_frame.pack(fill=tk.X)
        summary_total_frame.columnconfigure(1, weight=1)
        
        self.summary_vars = {'total': tk.StringVar(value=f"0.00 {self.currency}"), 'grand_total': tk.StringVar(value=f"0.00 {self.currency}"), 'overall_discount': tk.DoubleVar(value=0.0), 'overall_markup': tk.DoubleVar(value=0.0)}
        self.summary_vars['overall_discount'].trace_add('write', self.update_project_summary); self.summary_vars['overall_markup'].trace_add('write', self.update_project_summary)
        r=0
        self.summary_total_label = ttk.Label(summary_total_frame, text=lang_manager.get("summary_total"), style='Header.TLabel')
        self.summary_total_label.grid(row=r, column=0, sticky='w')
        ttk.Label(summary_total_frame, textvariable=self.summary_vars['total'], style='Header.TLabel').grid(row=r, column=1, sticky='e'); r+=1
        r = self._create_entry_row(summary_total_frame, "overall_discount", 0.0, r, self.summary_vars); r+=1
        r = self._create_entry_row(summary_total_frame, "overall_markup", 0.0, r, self.summary_vars); r+=1
        ttk.Separator(summary_total_frame, orient='horizontal').grid(row=r, column=0, columnspan=2, sticky='ew', pady=10); r+=1
        self.summary_grand_total_label = ttk.Label(summary_total_frame, text=lang_manager.get("summary_grand_total"), style='Total.TLabel')
        self.summary_grand_total_label.grid(row=r, column=0, sticky='w', pady=(5,0))
        ttk.Label(summary_total_frame, textvariable=self.summary_vars['grand_total'], style='Total.TLabel').grid(row=r, column=1, sticky='e', pady=(5,0)); r+=1
        self.generate_invoice_btn = ttk.Button(summary_total_frame, text=lang_manager.get("summary_generate_invoice"), command=self.open_invoice_preview)
        self.generate_invoice_btn.grid(row=r, column=0, columnspan=2, sticky="ew", pady=(15,0))

    def update_summary_tree_headings(self):
        self.summary_tree.heading('model', text=lang_manager.get("summary_col_model"))
        self.summary_tree.heading('price', text=lang_manager.get("summary_col_price"))
        self.summary_tree.column('price', anchor='e', width=100)

    def open_invoice_preview(self):
        if not LIBRARIES_INSTALLED:
            messagebox.showerror("Function Unavailable", "Required libraries for PDF creation are missing. Please see startup message.")
            return
        InvoicePreviewWindow(self)
    
    def update_project_summary(self, *args):
        for i in self.summary_tree.get_children(): self.summary_tree.delete(i)
        total_price = 0.0
        for tab_id in self.notebook.tabs():
            tab_widget = self.nametowidget(tab_id)
            tab_name = self.notebook.tab(tab_id, "text")
            price = getattr(tab_widget, 'final_price', 0.0)
            self.summary_tree.insert("", "end", values=(tab_name, f"{price:.2f} {self.currency}"))
            total_price += price
        discount = self.summary_vars['overall_discount'].get(); markup = self.summary_vars['overall_markup'].get()
        if discount > 100: discount = 100
        grand_total = total_price * (1 - discount / 100) * (1 + markup / 100)
        self.summary_vars['total'].set(f"{total_price:.2f} {self.currency}"); self.summary_vars['grand_total'].set(f"{grand_total:.2f} {self.currency}")
    
    def add_new_tab(self, title=None, data=None):
        if title is None: title = f"Model {len(self.notebook.tabs()) + 1}"
        tab = CalculationTab(self.notebook, self, style='TFrame')
        self.notebook.add(tab, text=title); self.notebook.select(tab)
        self.populate_all_presets(tab)
        if data: tab.set_data(data)
        self.update_project_summary()
        return tab

    def delete_current_tab(self):
        if len(self.notebook.tabs()) <= 1: messagebox.showwarning("Action Impossible", "Cannot delete the last tab."); return
        selected_tab = self.get_active_tab()
        if messagebox.askyesno("Confirm", f"Delete tab '{self.notebook.tab(selected_tab, 'text')}'?"):
            self.notebook.forget(selected_tab); self.update_project_summary()
    
    def rename_current_tab(self):
        if not self.notebook.tabs(): return
        selected_tab = self.get_active_tab(); old_name = self.notebook.tab(selected_tab, 'text')
        new_name = simpledialog.askstring("Rename", "Enter new name for the tab:", initialvalue=old_name)
        if new_name: self.notebook.tab(selected_tab, text=new_name); self.update_project_summary()

    def get_active_tab(self):
        if not self.notebook.tabs(): return None
        try: return self.nametowidget(self.notebook.select())
        except tk.TclError: return None

    def new_project(self):
        if messagebox.askyesno("New Project", "All unsaved changes will be lost. Create new project?"):
            while len(self.notebook.tabs()) > 0: self.notebook.forget(self.notebook.tabs()[0])
            self.add_new_tab(); self.summary_vars['overall_discount'].set(0); self.summary_vars['overall_markup'].set(0)

    def save_project(self):
        filepath = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Calculator Projects", "*.json")])
        if not filepath: return False
        project_data = [self.nametowidget(tab_id).get_data() for tab_id in self.notebook.tabs()]
        full_data = {'version': 4.1, 'tabs': project_data, 'summary': {k: v.get() for k, v in self.summary_vars.items() if isinstance(v, tk.DoubleVar)}}
        try:
            with open(filepath, 'w', encoding='utf-8') as f: json.dump(full_data, f, indent=4, ensure_ascii=False)
            messagebox.showinfo("Success", "Project saved successfully.")
            return True
        except Exception as e: 
            messagebox.showerror("Error", f"Failed to save project: {e}")
            return False

    def load_project(self):
        filepath = filedialog.askopenfilename(filetypes=[("Calculator Projects", "*.json")])
        if not filepath: return
        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f: data = json.load(f)
            if messagebox.askyesno("Open Project", "All current tabs will be closed. Continue?"):
                while len(self.notebook.tabs()) > 0: self.notebook.forget(self.notebook.tabs()[0])
                for tab_data in data.get('tabs', []): self.add_new_tab(tab_data.pop('tab_title', None), tab_data)
                if not self.notebook.tabs(): self.add_new_tab()
                summary_data = data.get('summary', {})
                for k, v in summary_data.items():
                    if k in self.summary_vars: self.summary_vars[k].set(v)
                self.update_project_summary()
        except Exception as e: messagebox.showerror("Error", f"Failed to open project: {e}")

    def create_section(self, parent, title_lang_key, category, tab_instance):
        if tab_instance not in self.sections:
            self.sections[tab_instance] = {}
        
        frame = ttk.Frame(parent, padding=(0, 0, 0, 15), style='TFrame')
        frame.pack(fill=tk.X, expand=True)
        header_label = ttk.Label(frame, text=lang_manager.get(title_lang_key), font=('Segoe UI', 13, 'bold'), style='Header.TLabel')
        header_label.pack(anchor="w", pady=(0, 10))
        
        if category in self.preset_fields: self._create_preset_controls(frame, category, tab_instance)
        
        fields_frame = ttk.Frame(frame, padding=(10, 5), style='TFrame', relief='solid', borderwidth=1)
        fields_frame.pack(fill=tk.X, expand=True)
        fields_frame.columnconfigure(0, weight=1); fields_frame.columnconfigure(1, weight=0)
        
        self.sections[tab_instance][category] = {'header_label': header_label, 'lang_key': title_lang_key, 'entry_rows': {}}

        extra_fields = {"plastic": OrderedDict([('total_weight_g', 100.0), ('support_percentage', 20.0), ('print_time_h', 5.0), ('electricity_cost_kwh', 0.3), ('printer_power_w', 250.0), ('depreciation_h', 0.5), ('failure_rate', 5.0)]), "adjustments": OrderedDict([('discount_percentage', 0.0), ('additional_markup_percentage', 0.0)])}
        row_counter = 0; current_fields = self.preset_fields.get(category, OrderedDict()).copy()
        if category in extra_fields: current_fields.update(extra_fields[category])
        
        for var_name, default_value in current_fields.items():
             self._create_entry_row(fields_frame, var_name, default_value, row_counter, tab_instance.variables, tab_instance, category); row_counter += 1

    def _create_entry_row(self, parent, var_name, default_value, row, var_dict, tab_instance=None, category=None):
        label_frame = ttk.Frame(parent, style='TFrame')
        label_frame.grid(row=row, column=0, sticky="w", padx=5, pady=5)
        
        label_lang_key = self.field_lang_keys.get(var_name, var_name)
        label = ttk.Label(label_frame, text=lang_manager.get(label_lang_key), style='TLabel')
        label.pack(side=tk.LEFT, anchor='w')
        
        help_text = self.help_texts.get(var_name)
        if help_text:
            help_button = ttk.Button(label_frame, text="?", style='Help.TButton', width=2)
            help_button.pack(side=tk.LEFT, anchor='w', padx=(5, 0))
            Tooltip(help_button, help_text, bg=self.tooltip_bg, fg=self.tooltip_fg)
            
        var = tk.DoubleVar(value=default_value)
        callback = self.update_project_summary if tab_instance is None else tab_instance.calculate_cost
        var.trace_add("write", callback)
        var_dict[var_name] = var
        entry = ttk.Entry(parent, textvariable=var, width=15); entry.grid(row=row, column=1, sticky="e", padx=5, pady=5)
        
        if tab_instance and category:
            self.sections[tab_instance][category]['entry_rows'][var_name] = {'label': label, 'help_button': help_button if help_text else None}
        
        return row + 1

    def _create_preset_controls(self, parent, category, tab_instance):
        controls_frame = ttk.Frame(parent, style='TFrame'); controls_frame.pack(fill=tk.X, expand=True, pady=(0, 5))
        controls_frame.columnconfigure(0, weight=1); combo = ttk.Combobox(controls_frame, state="readonly", width=30, font=('Segoe UI', 9))
        combo.grid(row=0, column=0, sticky="ew"); combo.bind("<<ComboboxSelected>>", lambda event, c=category: self.apply_preset(c))
        tab_instance.variables[f'{category}_combo'] = combo
        btn_frame = ttk.Frame(controls_frame, style='TFrame'); btn_frame.grid(row=0, column=1, sticky="e", padx=(10, 0))
        ttk.Button(btn_frame, text=lang_manager.get("preset_save"), command=lambda c=category: self.save_preset(c), style='Small.TButton').pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text=lang_manager.get("preset_edit"), command=lambda c=category: self.edit_preset(c), style='Small.TButton').pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text=lang_manager.get("preset_delete"), command=lambda c=category: self.delete_preset(c), style='Small.TButton').pack(side=tk.LEFT, padx=2)
    
    def _create_result_row(self, parent, label_lang_key, row, tab_instance, key_style='ResultKey.TLabel', value_style='ResultValue.TLabel'):
        label_widget = ttk.Label(parent, text=lang_manager.get(label_lang_key), style=key_style)
        label_widget.grid(row=row, column=0, sticky="w", padx=0, pady=2)
        
        var_name = label_lang_key
        var = tk.StringVar(value=f"0.00 {self.currency}"); 
        tab_instance.variables[var_name] = var
        tab_instance.result_labels[label_lang_key] = label_widget

        ttk.Label(parent, textvariable=var, style=value_style).grid(row=row, column=1, sticky="e", padx=0, pady=2)
        return row + 1
    
    def populate_all_presets(self, tab_instance):
        for category in self.preset_fields: self.populate_presets(category, tab_instance)

    def populate_presets(self, category, tab_instance):
        combo = tab_instance.variables.get(f'{category}_combo');
        if not combo: return
        preset_names = list(self.presets.get(category, {}).keys()); combo['values'] = preset_names
        if preset_names: combo.set(preset_names[0]); self.apply_preset(category, True)

    def apply_preset(self, category, force_apply=False):
        tab = self.get_active_tab()
        if not tab: return
        combo = tab.variables.get(f'{category}_combo')
        if not combo: return
        preset_name = combo.get()
        if preset_name in self.presets.get(category, {}):
            preset_data = self.presets[category][preset_name]
            for field, value in preset_data.items():
                if field in tab.variables: tab.variables[field].set(value)
        if not force_apply: tab.calculate_cost()

    def save_preset(self, category):
        tab = self.get_active_tab()
        if not tab: return
        preset_name = simpledialog.askstring("Save Preset", f"Name for '{category}' preset:")
        if not preset_name: return
        preset_data = {field: tab.variables[field].get() for field in self.preset_fields[category]}
        self.presets.setdefault(category, {})[preset_name] = preset_data; self.save_presets_to_file()
        for t_id in self.notebook.tabs(): self.populate_presets(category, self.nametowidget(t_id))
        tab.variables[f'{category}_combo'].set(preset_name)

    def edit_preset(self, category):
        tab = self.get_active_tab(); combo = tab.variables.get(f'{category}_combo')
        if not tab or not combo or not combo.get(): return
        old_name = combo.get()
        new_name = simpledialog.askstring("Edit Preset", "New name:", initialvalue=old_name)
        if not new_name: return
        preset_data = {field: tab.variables[field].get() for field in self.preset_fields[category]}
        if old_name != new_name and old_name in self.presets[category]: del self.presets[category][old_name]
        self.presets[category][new_name] = preset_data; self.save_presets_to_file()
        for t_id in self.notebook.tabs():
            current_tab = self.nametowidget(t_id); current_combo = current_tab.variables.get(f'{category}_combo')
            if current_combo:
                current_selection = current_combo.get()
                self.populate_presets(category, current_tab)
                if current_selection == old_name: current_combo.set(new_name)
                else: current_combo.set(current_selection)

    def delete_preset(self, category):
        tab = self.get_active_tab(); combo = tab.variables.get(f'{category}_combo')
        if not tab or not combo or not combo.get(): return
        preset_name = combo.get()
        if messagebox.askyesno("Delete Preset", f"Delete preset '{preset_name}'?"):
            if preset_name in self.presets.get(category, {}):
                del self.presets[category][preset_name]; self.save_presets_to_file()
                for t_id in self.notebook.tabs(): self.populate_presets(category, self.nametowidget(t_id))
    
    def generate_invoice_pdf(self, filepath, selected_models, lang_code, model_detail_settings, notes=""):
        c = pdfcanvas.Canvas(filepath, pagesize=A4)
        width, height = A4
        margin = 1.5 * cm
        y_pos, page_num = height - margin, 1
        
        font_name = "DejaVu" if os.path.exists('DejaVuSans.ttf') else "Helvetica"
        font_name_bold = "DejaVu-Bold" if os.path.exists('DejaVuSans-Bold.ttf') else "Helvetica-Bold"
       
        def draw_header():
            c.setFont(font_name, 14)
            c.drawString(margin, height - margin, self.invoice_settings.get('shop_name', 'Company Name'))
            logo_path = self.invoice_settings.get('logo_path', '')
            if logo_path and os.path.exists(logo_path):
                try: c.drawImage(logo_path, width - margin - (1.5*cm), height - margin - (0.5*cm), width=1.5*cm, height=1.5*cm, preserveAspectRatio=True, anchor='n')
                except: pass
            c.setStrokeColorRGB(0.8, 0.8, 0.8)
            c.line(margin, height - margin - (0.6*cm), width - margin, height - margin - (0.6*cm))
            return height - margin - (1.6*cm)

        def draw_footer():
            c.setFont(font_name, 7)
            c.setFillColorRGB(0.5, 0.5, 0.5)
            c.drawCentredString(width/2, margin/2 + 8, self.invoice_settings.get('footer_text', 'Thank you!'))
            c.drawRightString(width - margin, margin/2, f"{lang_manager.get('pdf_page', lang_code)} {page_num}")

        y_pos = draw_header(); draw_footer()

        for model_name in selected_models:
            price_str = ""
            for item in self.summary_tree.get_children():
                if self.summary_tree.item(item, 'values')[0] == model_name:
                    price_str = self.summary_tree.item(item, 'values')[1]
                    break
            
            tab_widget = self.nametowidget(next(t_id for t_id in self.notebook.tabs() if self.notebook.tab(t_id, "text") == model_name))
            settings = model_detail_settings.get(model_name, {})

            # Estimate block height
            est_height = 2 * cm
            for group_data in self.invoice_detail_groups.values():
                if any(settings.get(k, True) for k in group_data['items']):
                     est_height += 0.6 * cm 
                     for item_key in group_data['items']:
                        if settings.get(item_key, True):
                            est_height += 0.45 * cm
            if y_pos - est_height < margin + cm:
                 c.showPage(); page_num += 1; y_pos = draw_header(); draw_footer()

            c.setFillColor(colors.HexColor("#333333"))
            c.rect(margin, y_pos - 0.7*cm, width - 2*margin, 0.7*cm, fill=1, stroke=0)
            c.setFillColor(colors.white); c.setFont(font_name_bold, 11)
            c.drawString(margin + 0.3*cm, y_pos - 0.5*cm, model_name)
            c.drawRightString(width - margin - 0.3*cm, y_pos - 0.5*cm, price_str)
            y_pos -= (1.2*cm)
            
            if not tab_widget: continue

            for group_key, group_data in self.invoice_detail_groups.items():
                has_visible_items = any(settings.get(k, True) for k in group_data['items'])
                if not has_visible_items and group_key != "group_final": continue

                c.setFont(font_name_bold, 9)
                c.setFillColor(colors.black)
                group_total_val_str = tab_widget.variables.get(group_data['total_key'], tk.StringVar(value="")).get()
                
                if group_data['total_key'] or group_key == "group_final":
                    if group_data['total_key'] and has_visible_items:
                        c.drawString(margin, y_pos, lang_manager.get(group_data['lang_key'], lang_code))
                        c.drawRightString(width-margin, y_pos, group_total_val_str)
                        y_pos -= 0.5*cm
                
                c.setFont(font_name, 8)
                
                for item_key, item_lang_key in group_data['items'].items():
                    if settings.get(item_key, True):
                        raw_val = tab_widget.variables[item_lang_key].get()
                        
                        # Formatting for technical params
                        display_val = str(raw_val)
                        if item_key == 'show_weight':
                            try: display_val = f"{float(raw_val):.0f} g"
                            except: pass
                        elif item_key == 'show_time':
                            try: display_val = f"{float(raw_val):.1f} h"
                            except: pass

                        label = lang_manager.get(item_lang_key, lang_code)
                        if item_key == 'show_weight': label = lang_manager.get('param_weight', lang_code)
                        if item_key == 'show_time': label = lang_manager.get('param_time', lang_code)

                        c.drawString(margin + 0.5*cm, y_pos, label)
                        c.drawRightString(width - margin, y_pos, display_val)
                        y_pos -= 0.45*cm
                
                y_pos -= 0.2*cm
        
        total, discount, markup, grand_total = self.summary_vars['total'].get(), self.summary_vars['overall_discount'].get(), self.summary_vars['overall_markup'].get(), self.summary_vars['grand_total'].get()
        
        if y_pos - (4*cm) < margin + cm:
            c.showPage(); page_num += 1; y_pos = draw_header(); draw_footer()
        
        # Отрисовка заметок (Notes) в PDF
        if notes:
            c.setFont(font_name_bold, 9)
            c.drawString(margin, y_pos, lang_manager.get("invoice_notes_label", lang_code))
            y_pos -= 0.5*cm
            text_obj = c.beginText(margin, y_pos)
            text_obj.setFont(font_name, 8)
            words = notes.split()
            line = ""
            for word in words:
                if c.stringWidth(line + word, font_name, 8) < (width - 2*margin):
                    line += word + " "
                else:
                    text_obj.textLine(line)
                    line = word + " "
                    y_pos -= 0.4*cm
            text_obj.textLine(line)
            c.drawText(text_obj)
            y_pos -= 1.0*cm

        y_pos -= 0.5*cm; c.setStrokeColorRGB(0.2, 0.2, 0.2); c.line(width/2, y_pos, width - margin, y_pos); y_pos -= 0.5*cm
        
        c.setFont(f"{font_name}", 9)
        c.drawRightString(width - margin - 4*cm, y_pos, f"{lang_manager.get('pdf_total_amount', lang_code)}:")
        c.drawRightString(width - margin, y_pos, total)
        y_pos -= 0.5*cm
        
        if discount > 0:
            c.drawRightString(width - margin - 4*cm, y_pos, f"{lang_manager.get('pdf_discount', lang_code)} ({discount}%):")
            c.drawRightString(width - margin, y_pos, f"- {(float(total.split(' ')[0]) * discount / 100):.2f} {self.currency}")
            y_pos -= 0.5*cm
        if markup > 0:
            c.drawRightString(width - margin - 4*cm, y_pos, f"+ {(float(total.split(' ')[0]) * (1 - discount/100) * markup / 100):.2f} {self.currency}")
            y_pos -= 0.5*cm

        c.setFont(f"{font_name_bold}", 12)
        c.drawRightString(width - margin - 4*cm, y_pos - 0.2*cm, f"{lang_manager.get('pdf_total_due', lang_code)}:")
        c.drawRightString(width - margin, y_pos - 0.2*cm, grand_total)
        
        try:
            c.setEncrypt("", "", canModify=0, canCopy=0, canAnnotate=0, canPrint=1)
        except TypeError:
            try:
                c.setEncrypt("", "")
            except TypeError:
                c.setEncrypt("")

        c.save()

if __name__ == "__main__":
    app = CostCalculatorApp()
    app.mainloop()
