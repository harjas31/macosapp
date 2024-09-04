import sys
import customtkinter as ctk
from tkinter import filedialog, messagebox, BooleanVar, simpledialog
import logging
import os
import threading
import datetime
import pickle
import webbrowser
import time
import tempfile
import csv
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Import the required modules
from amazon_scraper import search as amazon_search
from flipkart_scraper import search as flipkart_search
from product_info_fetcher import fetch_amazon_product_info, fetch_flipkart_product_info
from export_utils import export_to_excel

class ProductInfoFetcherApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.title("Product Info Fetcher")
        self.geometry("800x700")

        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(expand=True, fill="both", padx=15, pady=15)

        # Create tabs
        self.amazon_rank_tab = self.tabview.add("Amazon Keyword Rank")
        self.flipkart_rank_tab = self.tabview.add("Flipkart Keyword Rank")
        self.amazon_product_tab = self.tabview.add("Amazon Product Info")
        self.flipkart_product_tab = self.tabview.add("Flipkart Product Info")

        self.create_amazon_rank_fetcher_tab(self.amazon_rank_tab)
        self.create_flipkart_rank_fetcher_tab(self.flipkart_rank_tab)
        self.create_amazon_product_info_tab(self.amazon_product_tab)
        self.create_flipkart_product_info_tab(self.flipkart_product_tab)

        # Initialize result dictionaries
        self.amazon_rank_results_checkbox = {}
        self.amazon_rank_results_other = {}
        self.flipkart_rank_results_checkbox = {}
        self.flipkart_rank_results_other = {}
        self.amazon_product_info_results = []
        self.flipkart_product_info_results = []

        # Initialize search status
        self.search_in_progress = False
        self.amazon_checkbox_search_completed = False
        self.amazon_other_search_completed = False
        self.flipkart_checkbox_search_completed = False
        self.flipkart_other_search_completed = False

        # Set up logging
        self.setup_logging()

        # Google Drive folder IDs configuration
        self.folder_ids = {
            'amazon': {
                'rank_fetcher': '142k5r7h-nAKFk2KadJWnCKHvApMKFpz7',
                'generic': '12zffS50CLv_yIbOU3ecuzjIOeCTaVEt5',
                'branded': '1mN5v26gHUcr4VBNZOY3irlge39ZeY7jv',
                'competition': '1SAe1-jmDLHXpSYOMdWbMm6qn7UL24PSK'
            },
            'flipkart': {
                'rank_fetcher': '1oolWrC8h1vMhg2VlFDdth8e2w7rM63Rq',
                'generic': '1O2OrRUQ4BT1Z_JqfbVk5jNXr4o1Zs6jV',
                'branded': '1b5ttslIw1D2Mkp1CAxPxpWtuUMCw63qC',
                'competition': '1bGJoAIFXiyO0oGTLSLMyEPbRT34ZToy4'
            }
        }

    def setup_logging(self):
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.INFO)
        handler = logging.StreamHandler(sys.stdout)
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        self.logger.addHandler(handler)

    def create_amazon_rank_fetcher_tab(self, parent):
        main_frame = ctk.CTkScrollableFrame(parent)
        main_frame.pack(expand=True, fill="both", padx=10, pady=10)

        # Checkbox section
        checkbox_frame = ctk.CTkFrame(main_frame)
        checkbox_frame.pack(fill="x", padx=5, pady=5)

        # Create three side-by-side frames for checkbox lists
        self.amazon_generic_vars = {}
        self.amazon_branded_vars = {}
        self.amazon_competition_vars = {}
        for i, (title, vars_dict, keywords) in enumerate([
            ("Generic", self.amazon_generic_vars, ["perfume", "perfume for men", "perfume for women", "unisex perfumes", "long lasting perfumes"]),
            ("Branded", self.amazon_branded_vars, ["bellavita perfumes", "bella vita luxury perfume for men", "bella vita perfume for women", "bella vita perfume for men"]),
            ("Competition", self.amazon_competition_vars, ["park avenue perfume for men", "wild stone perfume for men", "renee perfume"])
        ]):
            column_frame = ctk.CTkFrame(checkbox_frame)
            column_frame.grid(row=0, column=i, padx=5, pady=5, sticky="nsew")
            checkbox_frame.grid_columnconfigure(i, weight=1)

            ctk.CTkLabel(column_frame, text=title, font=("Arial", 14, "bold")).pack(pady=(0, 5))
            for keyword in keywords:
                var = ctk.BooleanVar()
                ctk.CTkCheckBox(column_frame, text=keyword, variable=var).pack(anchor="w", pady=2)
                vars_dict[keyword] = var

        # Ranking entry for checkbox section
        ranking_frame = ctk.CTkFrame(main_frame)
        ranking_frame.pack(fill="x", padx=5, pady=5)
        ctk.CTkLabel(ranking_frame, text="Enter ranking (up to 100):").pack(side="left", padx=(0, 5))
        self.amazon_rank_entry_checkbox = ctk.CTkEntry(ranking_frame, width=100)
        self.amazon_rank_entry_checkbox.pack(side="left")

        # Buttons for checkbox section
        button_frame = ctk.CTkFrame(main_frame)
        button_frame.pack(fill="x", padx=5, pady=5)
        self.amazon_rank_checkbox_button = ctk.CTkButton(button_frame, text="Fetch Selected Keywords", command=self.process_amazon_rank_fetcher_checkboxes)
        self.amazon_rank_checkbox_button.pack(side="left", padx=(0, 5), expand=True, fill="x")
        self.amazon_rank_save_cloud_button_checkbox = ctk.CTkButton(button_frame, text="Save to Cloud", command=self.save_amazon_rank_to_cloud_checkbox)
        self.amazon_rank_save_cloud_button_checkbox.pack(side="left", expand=True, fill="x")
        self.amazon_rank_save_cloud_button_checkbox.configure(state="disabled")

        # Status label for checkbox section
        self.amazon_rank_status_checkbox = ctk.CTkLabel(main_frame, text="Status: Ready", text_color="white")
        self.amazon_rank_status_checkbox.pack(pady=5)

        # Separator
        ctk.CTkFrame(main_frame, height=2, fg_color="gray").pack(fill="x", padx=5, pady=10)

        # Other keywords section
        ctk.CTkLabel(main_frame, text="Other Keywords (one per line):", font=("Arial", 14, "bold")).pack(anchor="w", padx=5, pady=5)
        self.amazon_rank_keywords = ctk.CTkTextbox(main_frame, height=100)
        self.amazon_rank_keywords.pack(fill="x", padx=5, pady=5)

        # Ranking entry for other keywords section
        other_ranking_frame = ctk.CTkFrame(main_frame)
        other_ranking_frame.pack(fill="x", padx=5, pady=5)
        ctk.CTkLabel(other_ranking_frame, text="Enter ranking (up to 100):").pack(side="left", padx=(0, 5))
        self.amazon_rank_entry_other = ctk.CTkEntry(other_ranking_frame, width=100)
        self.amazon_rank_entry_other.pack(side="left")

        # Buttons for other keywords section
        other_button_frame = ctk.CTkFrame(main_frame)
        other_button_frame.pack(fill="x", padx=5, pady=5)
        self.amazon_rank_other_button = ctk.CTkButton(other_button_frame, text="Fetch Other Keywords", command=self.process_amazon_rank_fetcher_other)
        self.amazon_rank_other_button.pack(side="left", padx=(0, 5), expand=True, fill="x")
        self.amazon_rank_save_cloud_button_other = ctk.CTkButton(other_button_frame, text="Save to Cloud", command=self.save_amazon_rank_to_cloud_other)
        self.amazon_rank_save_cloud_button_other.pack(side="left", expand=True, fill="x")
        self.amazon_rank_save_cloud_button_other.configure(state="disabled")

        # Status label for other keywords section
        self.amazon_rank_status_other = ctk.CTkLabel(main_frame, text="Status: Ready", text_color="white")
        self.amazon_rank_status_other.pack(pady=5)

    def create_flipkart_rank_fetcher_tab(self, parent):
        main_frame = ctk.CTkScrollableFrame(parent)
        main_frame.pack(expand=True, fill="both", padx=10, pady=10)

        # Checkbox section
        checkbox_frame = ctk.CTkFrame(main_frame)
        checkbox_frame.pack(fill="x", padx=5, pady=5)

        # Create three side-by-side frames for checkbox lists
        self.flipkart_generic_vars = {}
        self.flipkart_branded_vars = {}
        self.flipkart_competition_vars = {}
        for i, (title, vars_dict, keywords) in enumerate([
            ("Generic", self.flipkart_generic_vars, ["perfume", "perfume for men", "perfume for women", "unisex perfumes", "long lasting perfumes"]),
            ("Branded", self.flipkart_branded_vars, ["bellavita perfumes", "bella vita luxury perfume for men", "bella vita perfume for women", "bella vita perfume for men"]),
            ("Competition", self.flipkart_competition_vars, ["park avenue perfume for men", "wild stone perfume for men", "renee perfume"])
        ]):
            column_frame = ctk.CTkFrame(checkbox_frame)
            column_frame.grid(row=0, column=i, padx=5, pady=5, sticky="nsew")
            checkbox_frame.grid_columnconfigure(i, weight=1)

            ctk.CTkLabel(column_frame, text=title, font=("Arial", 14, "bold")).pack(pady=(0, 5))
            for keyword in keywords:
                var = ctk.BooleanVar()
                ctk.CTkCheckBox(column_frame, text=keyword, variable=var).pack(anchor="w", pady=2)
                vars_dict[keyword] = var

        # Ranking entry for checkbox section
        ranking_frame = ctk.CTkFrame(main_frame)
        ranking_frame.pack(fill="x", padx=5, pady=5)
        ctk.CTkLabel(ranking_frame, text="Enter ranking (up to 100):").pack(side="left", padx=(0, 5))
        self.flipkart_rank_entry_checkbox = ctk.CTkEntry(ranking_frame, width=100)
        self.flipkart_rank_entry_checkbox.pack(side="left")

        # Buttons for checkbox section
        button_frame = ctk.CTkFrame(main_frame)
        button_frame.pack(fill="x", padx=5, pady=5)
        self.flipkart_rank_checkbox_button = ctk.CTkButton(button_frame, text="Fetch Selected Keywords", command=self.process_flipkart_rank_fetcher_checkboxes)
        self.flipkart_rank_checkbox_button.pack(side="left", padx=(0, 5), expand=True, fill="x")
        self.flipkart_rank_save_cloud_button_checkbox = ctk.CTkButton(button_frame, text="Save to Cloud", command=self.save_flipkart_rank_to_cloud_checkbox)
        self.flipkart_rank_save_cloud_button_checkbox.pack(side="left", expand=True, fill="x")
        self.flipkart_rank_save_cloud_button_checkbox.configure(state="disabled")

        # Status label for checkbox section
        self.flipkart_rank_status_checkbox = ctk.CTkLabel(main_frame, text="Status: Ready", text_color="white")
        self.flipkart_rank_status_checkbox.pack(pady=5)

        # Separator
        ctk.CTkFrame(main_frame, height=2, fg_color="gray").pack(fill="x", padx=5, pady=10)

        # Other keywords section
        ctk.CTkLabel(main_frame, text="Other Keywords (one per line):", font=("Arial", 14, "bold")).pack(anchor="w", padx=5, pady=5)
        self.flipkart_rank_keywords = ctk.CTkTextbox(main_frame, height=100)
        self.flipkart_rank_keywords.pack(fill="x", padx=5, pady=5)

        # Ranking entry for other keywords section
        other_ranking_frame = ctk.CTkFrame(main_frame)
        other_ranking_frame.pack(fill="x", padx=5, pady=5)
        ctk.CTkLabel(other_ranking_frame, text="Enter ranking (up to 100):").pack(side="left", padx=(0, 5))
        self.flipkart_rank_entry_other = ctk.CTkEntry(other_ranking_frame, width=100)
        self.flipkart_rank_entry_other.pack(side="left")

        # Buttons for other keywords section
        other_button_frame = ctk.CTkFrame(main_frame)
        other_button_frame.pack(fill="x", padx=5, pady=5)
        self.flipkart_rank_other_button = ctk.CTkButton(other_button_frame, text="Fetch Other Keywords", command=self.process_flipkart_rank_fetcher_other)
        self.flipkart_rank_other_button.pack(side="left", padx=(0, 5), expand=True, fill="x")
        self.flipkart_rank_save_cloud_button_other = ctk.CTkButton(other_button_frame, text="Save to Cloud", command=self.save_flipkart_rank_to_cloud_other)
        self.flipkart_rank_save_cloud_button_other.pack(side="left", expand=True, fill="x")
        self.flipkart_rank_save_cloud_button_other.configure(state="disabled")

        # Status label for other keywords section
        self.flipkart_rank_status_other = ctk.CTkLabel(main_frame, text="Status: Ready", text_color="white")
        self.flipkart_rank_status_other.pack(pady=5)

    def create_amazon_product_info_tab(self, parent):
        frame = ctk.CTkFrame(parent)
        frame.pack(expand=True, fill="both", padx=10, pady=10)

        ctk.CTkLabel(frame, text="Enter Link(s) / ASIN(s) (one per line):").pack(padx=10, pady=5)
        self.amazon_product_links = ctk.CTkTextbox(frame, height=100)
        self.amazon_product_links.pack(padx=10, pady=5, fill="x")

        self.amazon_product_button = ctk.CTkButton(frame, text="Fetch Amazon Product Info", command=self.process_amazon_product_info)
        self.amazon_product_button.pack(padx=10, pady=10)

        self.amazon_product_status = ctk.CTkLabel(frame, text="Status: Ready", text_color="white")
        self.amazon_product_status.pack(padx=10, pady=5)

    def create_flipkart_product_info_tab(self, parent):
        frame = ctk.CTkFrame(parent)
        frame.pack(expand=True, fill="both", padx=10, pady=10)

        ctk.CTkLabel(frame, text="Enter Link(s) (one per line):").pack(padx=10, pady=5)
        self.flipkart_product_links = ctk.CTkTextbox(frame, height=100)
        self.flipkart_product_links.pack(padx=10, pady=5, fill="x")

        self.flipkart_product_button = ctk.CTkButton(frame, text="Fetch Flipkart Product Info", command=self.process_flipkart_product_info)
        self.flipkart_product_button.pack(padx=10, pady=10)

        self.flipkart_product_status = ctk.CTkLabel(frame, text="Status: Ready", text_color="white")
        self.flipkart_product_status.pack(padx=10, pady=5)

    def process_amazon_rank_fetcher_checkboxes(self):
        if self.search_in_progress:
            messagebox.showerror("Error", "A search is already in progress.")
            return

        selected_keywords = []
        for keyword_list in [self.amazon_generic_vars, self.amazon_branded_vars, self.amazon_competition_vars]:
            selected_keywords.extend([k for k, v in keyword_list.items() if v.get()])

        if not selected_keywords:
            messagebox.showerror("Error", "Please select at least one keyword.")
            return

        ranking = self.amazon_rank_entry_checkbox.get()
        if not ranking.isdigit() or int(ranking) > 100:
            messagebox.showerror("Invalid Input", "Please enter a ranking number up to 100.")
            return

        self.search_in_progress = True
        self.amazon_rank_results_checkbox = {}  # Clear previous results
        self.amazon_rank_status_checkbox.configure(text=f"Status: Processing (0/{len(selected_keywords)})", text_color="white")
        self.amazon_rank_checkbox_button.configure(state="disabled")
        self.amazon_rank_other_button.configure(state="disabled")

        threading.Thread(target=self._process_amazon_rank_fetcher, args=(selected_keywords, int(ranking), "checkbox")).start()

    def process_amazon_rank_fetcher_other(self):
        if self.search_in_progress:
            messagebox.showerror("Error", "A search is already in progress.")
            return

        keywords = self.amazon_rank_keywords.get("1.0", "end-1c").splitlines()
        ranking = self.amazon_rank_entry_other.get()

        if not keywords or not ranking.isdigit() or int(ranking) > 100:
            messagebox.showerror("Invalid Input", "Please enter at least one keyword and a ranking number up to 100.")
            return

        self.search_in_progress = True
        self.amazon_rank_results_other = {}  # Clear previous results
        self.amazon_rank_status_other.configure(text=f"Status: Processing (0/{len(keywords)})", text_color="white")
        self.amazon_rank_checkbox_button.configure(state="disabled")
        self.amazon_rank_other_button.configure(state="disabled")

        threading.Thread(target=self._process_amazon_rank_fetcher, args=(keywords, int(ranking), "other")).start()

    def _process_amazon_rank_fetcher(self, keywords, ranking, section):
        try:
            results = {}
            for i, keyword in enumerate(keywords, 1):
                result = amazon_search(keyword, ranking)
                results[keyword] = result
                if section == "checkbox":
                    self.amazon_rank_status_checkbox.configure(text=f"Status: Processing ({i}/{len(keywords)})", text_color="white")
                else:
                    self.amazon_rank_status_other.configure(text=f"Status: Processing ({i}/{len(keywords)})", text_color="white")
            
            if section == "checkbox":
                self.amazon_rank_results_checkbox = results
                self.amazon_rank_save_cloud_button_checkbox.configure(state="normal")
                self.amazon_checkbox_search_completed = True
                self.amazon_rank_status_checkbox.configure(text="Status: Completed", text_color="green")
            else:
                self.amazon_rank_results_other = results
                self.amazon_rank_save_cloud_button_other.configure(state="normal")
                self.amazon_other_search_completed = True
                self.amazon_rank_status_other.configure(text="Status: Completed", text_color="green")
            
            self.save_results(results, 'Amazon Rank Fetcher')
        except Exception as e:
            self.logger.error(f"Error in Amazon Rank Fetcher: {str(e)}")
            if section == "checkbox":
                self.amazon_rank_status_checkbox.configure(text="Status: Error occurred", text_color="red")
            else:
                self.amazon_rank_status_other.configure(text="Status: Error occurred", text_color="red")
        finally:
            self.search_in_progress = False
            self.amazon_rank_checkbox_button.configure(state="normal")
            self.amazon_rank_other_button.configure(state="normal")

    def process_flipkart_rank_fetcher_checkboxes(self):
        if self.search_in_progress:
            messagebox.showerror("Error", "A search is already in progress.")
            return

        selected_keywords = []
        for keyword_list in [self.flipkart_generic_vars, self.flipkart_branded_vars, self.flipkart_competition_vars]:
            selected_keywords.extend([k for k, v in keyword_list.items() if v.get()])

        if not selected_keywords:
            messagebox.showerror("Error", "Please select at least one keyword.")
            return

        ranking = self.flipkart_rank_entry_checkbox.get()
        if not ranking.isdigit() or int(ranking) > 100:
            messagebox.showerror("Invalid Input", "Please enter a ranking number up to 100.")
            return

        self.search_in_progress = True
        self.flipkart_rank_results_checkbox = {}  # Clear previous results
        self.flipkart_rank_status_checkbox.configure(text=f"Status: Processing (0/{len(selected_keywords)})", text_color="white")
        self.flipkart_rank_checkbox_button.configure(state="disabled")
        self.flipkart_rank_other_button.configure(state="disabled")

        threading.Thread(target=self._process_flipkart_rank_fetcher, args=(selected_keywords, int(ranking), "checkbox")).start()

    def process_flipkart_rank_fetcher_other(self):
        if self.search_in_progress:
            messagebox.showerror("Error", "A search is already in progress.")
            return

        keywords = self.flipkart_rank_keywords.get("1.0", "end-1c").splitlines()
        ranking = self.flipkart_rank_entry_other.get()

        if not keywords or not ranking.isdigit() or int(ranking) > 100:
            messagebox.showerror("Invalid Input", "Please enter at least one keyword and a ranking number up to 100.")
            return

        self.search_in_progress = True
        self.flipkart_rank_results_other = {}  # Clear previous results
        self.flipkart_rank_status_other.configure(text=f"Status: Processing (0/{len(keywords)})", text_color="white")
        self.flipkart_rank_checkbox_button.configure(state="disabled")
        self.flipkart_rank_other_button.configure(state="disabled")

        threading.Thread(target=self._process_flipkart_rank_fetcher, args=(keywords, int(ranking), "other")).start()

    def _process_flipkart_rank_fetcher(self, keywords, ranking, section):
        try:
            results = {}
            for i, keyword in enumerate(keywords, 1):
                result = flipkart_search(keyword, ranking)
                results[keyword] = result
                if section == "checkbox":
                    self.flipkart_rank_status_checkbox.configure(text=f"Status: Processing ({i}/{len(keywords)})", text_color="white")
                else:
                    self.flipkart_rank_status_other.configure(text=f"Status: Processing ({i}/{len(keywords)})", text_color="white")
            
            if section == "checkbox":
                self.flipkart_rank_results_checkbox = results
                self.flipkart_rank_save_cloud_button_checkbox.configure(state="normal")
                self.flipkart_checkbox_search_completed = True
                self.flipkart_rank_status_checkbox.configure(text="Status: Completed", text_color="green")
            else:
                self.flipkart_rank_results_other = results
                self.flipkart_rank_save_cloud_button_other.configure(state="normal")
                self.flipkart_other_search_completed = True
                self.flipkart_rank_status_other.configure(text="Status: Completed", text_color="green")
            
            self.save_results(results, 'Flipkart Rank Fetcher')
        except Exception as e:
            self.logger.error(f"Error in Flipkart Rank Fetcher: {str(e)}")
            if section == "checkbox":
                self.flipkart_rank_status_checkbox.configure(text="Status: Error occurred", text_color="red")
            else:
                self.flipkart_rank_status_other.configure(text="Status: Error occurred", text_color="red")
        finally:
            self.search_in_progress = False
            self.flipkart_rank_checkbox_button.configure(state="normal")
            self.flipkart_rank_other_button.configure(state="normal")

    def process_amazon_product_info(self):
        links = self.amazon_product_links.get("1.0", "end-1c").splitlines()

        if not links:
            messagebox.showerror("Invalid Input", "Please enter at least one link or ASIN.")
            return

        if self.search_in_progress:
            messagebox.showerror("Error", "A search is already in progress.")
            return

        self.search_in_progress = True
        self.amazon_product_info_results = []  # Clear previous results
        self.amazon_product_status.configure(text=f"Status: Processing (0/{len(links)})", text_color="white")
        self.amazon_product_button.configure(state="disabled")

        threading.Thread(target=self._process_amazon_product_info, args=(links,)).start()

    def _process_amazon_product_info(self, links):
        try:
            for i, link in enumerate(links, 1):
                result = fetch_amazon_product_info(link)
                if result:
                    self.amazon_product_info_results.append(result)
                self.amazon_product_status.configure(text=f"Status: Processing ({i}/{len(links)})", text_color="white")
            
            self.save_results({'product': self.amazon_product_info_results}, 'Amazon Product Info')
            self.amazon_product_status.configure(text="Status: Completed", text_color="green")
        except Exception as e:
            self.logger.error(f"Error in Amazon Product Info: {str(e)}")
            self.amazon_product_status.configure(text="Status: Error occurred", text_color="red")
        finally:
            self.search_in_progress = False
            self.amazon_product_button.configure(state="normal")

    def process_flipkart_product_info(self):
        links = self.flipkart_product_links.get("1.0", "end-1c").splitlines()

        if not links:
            messagebox.showerror("Invalid Input", "Please enter at least one link.")
            return

        if self.search_in_progress:
            messagebox.showerror("Error", "A search is already in progress.")
            return

        self.search_in_progress = True
        self.flipkart_product_info_results = []  # Clear previous results
        self.flipkart_product_status.configure(text=f"Status: Processing (0/{len(links)})", text_color="white")
        self.flipkart_product_button.configure(state="disabled")

        threading.Thread(target=self._process_flipkart_product_info, args=(links,)).start()
    
    def _process_flipkart_product_info(self, links):
        try:
            for i, link in enumerate(links, 1):
                result = fetch_flipkart_product_info(link)
                if result:
                    self.flipkart_product_info_results.append(result)
                self.flipkart_product_status.configure(text=f"Status: Processing ({i}/{len(links)})", text_color="white")
            
            self.save_results({'product': self.flipkart_product_info_results}, 'Flipkart Product Info')
            self.flipkart_product_status.configure(text="Status: Completed", text_color="green")
        except Exception as e:
            self.logger.error(f"Error in Flipkart Product Info: {str(e)}")
            self.flipkart_product_status.configure(text="Status: Error occurred", text_color="red")
        finally:
            self.search_in_progress = False
            self.flipkart_product_button.configure(state="normal")

    def save_amazon_rank_to_cloud_checkbox(self):
        if not self.amazon_checkbox_search_completed:
            messagebox.showwarning("No Results", "There are no new results to save to the cloud.")
            return
        self._save_rank_to_cloud(self.amazon_rank_results_checkbox, "Amazon", "checkbox")

    def save_amazon_rank_to_cloud_other(self):
        if not self.amazon_other_search_completed:
            messagebox.showwarning("No Results", "There are no new results to save to the cloud.")
            return
        self._save_rank_to_cloud(self.amazon_rank_results_other, "Amazon", "other")

    def save_flipkart_rank_to_cloud_checkbox(self):
        if not self.flipkart_checkbox_search_completed:
            messagebox.showwarning("No Results", "There are no new results to save to the cloud.")
            return
        self._save_rank_to_cloud(self.flipkart_rank_results_checkbox, "Flipkart", "checkbox")

    def save_flipkart_rank_to_cloud_other(self):
        if not self.flipkart_other_search_completed:
            messagebox.showwarning("No Results", "There are no new results to save to the cloud.")
            return
        self._save_rank_to_cloud(self.flipkart_rank_results_other, "Flipkart", "other")

    def _save_rank_to_cloud(self, results, platform, section):
        if not results:
            messagebox.showwarning("No Results", "There are no results to save to the cloud.")
            return

        def save_thread():
            try:
                drive_service = self.get_google_drive_service()
                if not drive_service:
                    return

                total_keywords = len(results)
                for i, (keyword, keyword_results) in enumerate(results.items(), 1):
                    folder_id = self.get_folder_id_for_keyword(keyword, platform)
                    file_name = f"{keyword} - {datetime.datetime.now().strftime('%H:%M - %d/%m/%y')}.csv"
                    self.save_to_drive_csv(drive_service, keyword_results, file_name, folder_id)
                    
                    if section == "checkbox":
                        if platform == "Amazon":
                            self.amazon_rank_status_checkbox.configure(text=f"Saving to cloud: {i}/{total_keywords}", text_color="blue")
                        else:
                            self.flipkart_rank_status_checkbox.configure(text=f"Saving to cloud: {i}/{total_keywords}", text_color="blue")
                    else:
                        if platform == "Amazon":
                            self.amazon_rank_status_other.configure(text=f"Saving to cloud: {i}/{total_keywords}", text_color="blue")
                        else:
                            self.flipkart_rank_status_other.configure(text=f"Saving to cloud: {i}/{total_keywords}", text_color="blue")

                messagebox.showinfo("Success", f"Results saved successfully to Google Drive for {platform}")
                
                if section == "checkbox":
                    if platform == "Amazon":
                        self.amazon_checkbox_search_completed = False
                        self.amazon_rank_status_checkbox.configure(text="Status: Saved to cloud", text_color="green")
                    else:
                        self.flipkart_checkbox_search_completed = False
                        self.flipkart_rank_status_checkbox.configure(text="Status: Saved to cloud", text_color="green")
                else:
                    if platform == "Amazon":
                        self.amazon_other_search_completed = False
                        self.amazon_rank_status_other.configure(text="Status: Saved to cloud", text_color="green")
                    else:
                        self.flipkart_other_search_completed = False
                        self.flipkart_rank_status_other.configure(text="Status: Saved to cloud", text_color="green")

            except Exception as e:
                self.logger.error(f"Error saving results to Google Drive: {str(e)}")
                messagebox.showerror("Error", f"Error saving results to Google Drive: {str(e)}")

        threading.Thread(target=save_thread).start()

    def get_folder_id_for_keyword(self, keyword, platform):
        if platform == "Amazon":
            if keyword in self.amazon_generic_vars:
                return self.folder_ids['amazon']['generic']
            elif keyword in self.amazon_branded_vars:
                return self.folder_ids['amazon']['branded']
            elif keyword in self.amazon_competition_vars:
                return self.folder_ids['amazon']['competition']
            else:
                return self.folder_ids['amazon']['rank_fetcher']
        elif platform == "Flipkart":
            if keyword in self.flipkart_generic_vars:
                return self.folder_ids['flipkart']['generic']
            elif keyword in self.flipkart_branded_vars:
                return self.folder_ids['flipkart']['branded']
            elif keyword in self.flipkart_competition_vars:
                return self.folder_ids['flipkart']['competition']
            else:
                return self.folder_ids['flipkart']['rank_fetcher']

    def save_to_drive_csv(self, service, results, file_name, folder_id):
        temp_file_name = None
        try:
            with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.csv', newline='', encoding='utf-8-sig') as temp_file:
                temp_file_name = temp_file.name
                writer = csv.DictWriter(temp_file, fieldnames=results[0].keys())
                writer.writeheader()
                for row in results:
                    cleaned_row = {k: str(v).encode('utf-8', errors='replace').decode('utf-8') for k, v in row.items()}
                    writer.writerow(cleaned_row)

            time.sleep(0.5)

            file_metadata = {'name': file_name, 'parents': [folder_id]}
            media = MediaFileUpload(temp_file_name, resumable=True, mimetype='text/csv')
            file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()

            self.logger.info(f"CSV file uploaded successfully to Google Drive. File ID: {file.get('id')}")

        except Exception as e:
            self.logger.error(f"Error in save_to_drive_csv: {str(e)}")
            raise
        finally:
            if temp_file_name:
                for _ in range(5):
                    try:
                        os.unlink(temp_file_name)
                        self.logger.info(f"Temporary CSV file {temp_file_name} deleted successfully")
                        break
                    except PermissionError:
                        time.sleep(1)
                else:
                    self.logger.warning(f"Failed to delete temporary file {temp_file_name}")

    def save_results(self, results, title):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"{title.replace(' ', '_')}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )

        if file_path:
            try:
                export_to_excel(results, file_path, title.split()[0])
                messagebox.showinfo("Success", f"Results saved successfully to {file_path}")
            except Exception as e:
                self.logger.error(f"Error saving results: {str(e)}")
                messagebox.showerror("Error", f"Error saving results: {str(e)}")

    def get_google_drive_service(self):
        creds = None
        SCOPES = ['https://www.googleapis.com/auth/drive.file']

        try:
            if os.path.exists('token.pickle'):
                with open('token.pickle', 'rb') as token:
                    creds = pickle.load(token)
            
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = Flow.from_client_secrets_file(
                        'client_secret.json', SCOPES)
                    flow.redirect_uri = 'urn:ietf:wg:oauth:2.0:oob'

                    auth_url, _ = flow.authorization_url(prompt='consent')
                    
                    webbrowser.open(auth_url)
                    
                    code = self.get_authorization_code()
                    flow.fetch_token(code=code)
                    creds = flow.credentials

                with open('token.pickle', 'wb') as token:
                    pickle.dump(creds, token)

            return build('drive', 'v3', credentials=creds)

        except Exception as e:
            self.logger.error(f"Error in Google Drive authentication: {str(e)}")
            messagebox.showerror("Authentication Error", f"Failed to authenticate with Google Drive: {str(e)}")
            return None

    def get_authorization_code(self):
        return simpledialog.askstring("Authorization Code", "Enter the authorization code:")

if __name__ == "__main__":
    app = ProductInfoFetcherApp()
    app.mainloop()