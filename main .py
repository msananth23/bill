import os
from datetime import datetime

from kivy.lang import Builder
from kivy.app import App
from kivy.properties import StringProperty, ListProperty
from kivymd.uix.screen import MDScreen
from kivymd.uix.list import OneLineListItem
from kivymd.uix.dialog import MDDialog
from kivymd.uix.button import MDFlatButton

import billing_core as core

Builder.load_file(os.path.join(os.path.dirname(__file__), "document_editor.kv"))

DEFAULT_TERMS = ["Goods once sold will not be taken back.", "Payment due within 15 days of invoice date."]


class DocumentEditorScreen(MDScreen):
    SCREEN_NAME = "document_editor"
    doc_type = StringProperty("Tax Invoice")
    items = ListProperty([])

    def on_enter_params(self, doc_type="Tax Invoice"):
        self.doc_type = doc_type
        self.items = []
        self.ids.company_name.text = core.DEFAULT_COMPANY["name"]
        self.ids.company_addr.text = core.DEFAULT_COMPANY["address"]
        self.ids.company_gstin.text = core.DEFAULT_COMPANY["gstin"]
        self.ids.company_contact.text = core.DEFAULT_COMPANY["contact"]
        self.ids.bank_name.text = core.DEFAULT_COMPANY["bank_name"]
        self.ids.bank_ac.text = core.DEFAULT_COMPANY["ac_no"]
        self.ids.bank_ifsc.text = core.DEFAULT_COMPANY["ifsc"]
        self.ids.buyer_name.text = ""
        self.ids.buyer_addr.text = ""
        self.ids.buyer_gstin.text = ""
        self.ids.inv_date.text = datetime.now().strftime("%d-%m-%Y")
        self.refresh_item_list()

    def go_back(self):
        App.get_running_app().go("dashboard")

    # -- item ledger ------------------------------------------------------------
    def add_item(self):
        desc = self.ids.item_desc.text.strip()
        hsn = self.ids.item_hsn.text.strip()
        qty = self.ids.item_qty.text.strip()
        rate = self.ids.item_rate.text.strip()
        gst_rate = self.ids.item_gst.text.strip() or "18"
        unit = self.ids.item_unit.text.strip() or "Nos"

        if not (desc and hsn and qty and rate):
            self._notify("Fill description, HSN, quantity and rate.")
            return
        try:
            qty_v = int(qty)
            rate_v = float(rate)
            gst_v = float(gst_rate)
        except ValueError:
            self._notify("Quantity, rate and GST% must be numbers.")
            return

        new_items = list(self.items)
        new_items.append({
            "SI No.": len(new_items) + 1, "Description of Goods": desc, "HSN/SAC": hsn,
            "Quantity": qty_v, "Unit": unit, "Unit Price": rate_v, "GST Type": "GST", "GST Rate": gst_v,
        })
        self.items = new_items
        self.ids.item_desc.text = ""
        self.ids.item_hsn.text = ""
        self.ids.item_qty.text = ""
        self.ids.item_rate.text = ""
        self.refresh_item_list()

    def remove_item(self, si_no):
        new_items = [i for i in self.items if i["SI No."] != si_no]
        for idx, i in enumerate(new_items):
            i["SI No."] = idx + 1
        self.items = new_items
        self.refresh_item_list()

    def refresh_item_list(self):
        box = self.ids.item_list_box
        box.clear_widgets()
        subtotal = 0.0
        for item in self.items:
            amount = item["Quantity"] * item["Unit Price"]
            subtotal += amount
            label = f"{item['SI No.']}. {item['Description of Goods']}  x{item['Quantity']} @ {item['Unit Price']} = Rs.{amount:.2f}"
            row = OneLineListItem(text=label, on_release=lambda w, sn=item["SI No."]: self.remove_item(sn))
            box.add_widget(row)
        self.ids.subtotal_label.text = f"Subtotal: Rs. {subtotal:.2f}  (tap a row to remove it)"

    # -- generate -----------------------------------------------------------------
    def generate_document(self):
        if not self.items:
            self._notify("Add at least one item first.")
            return
        company_info = {
            "name": self.ids.company_name.text.strip(), "address": self.ids.company_addr.text.strip(),
            "gstin": self.ids.company_gstin.text.strip(), "contact": self.ids.company_contact.text.strip(),
            "email": "", "logo_path": "",
        }
        buyer_info = {
            "name": self.ids.buyer_name.text.strip(), "address": self.ids.buyer_addr.text.strip(),
            "gstin": self.ids.buyer_gstin.text.strip(), "contact_name": "", "mobile": "",
        }
        bank_info = {
            "bank_name": self.ids.bank_name.text.strip(), "ac_no": self.ids.bank_ac.text.strip(),
            "ifsc": self.ids.bank_ifsc.text.strip(),
        }
        if not company_info["name"] or not buyer_info["name"] or not buyer_info["address"]:
            self._notify("Company name, buyer name and buyer address are required.")
            return

        app = App.get_running_app()
        formatted, number, counter_file, confirmed = core.get_next_document_number(
            self.doc_type, company_info["name"], sync_client=app.sync_client if app.sync_client.is_online() else None
        )
        core.save_document_number(counter_file, number)

        meta_info = {
            "invoice_no": formatted, "invoice_date": self.ids.inv_date.text.strip(),
            "po_number": "", "po_date": "",
        }

        try:
            pdf_path = core.generate_pdf(self.doc_type, company_info, buyer_info, bank_info, DEFAULT_TERMS, meta_info, self.items)
            xlsx_path = core.generate_excel(self.doc_type, company_info, buyer_info, bank_info, DEFAULT_TERMS, meta_info, self.items)
            core.update_summary_excel(company_info, self.doc_type, meta_info, buyer_info, self.items)
            core.update_hsn_summary_excel(company_info, meta_info, self.items)
        except Exception as e:
            self._notify(f"Generation failed: {e}")
            return

        state = {
            "doc_type": self.doc_type, "invoice_date": meta_info["invoice_date"],
            "company_info": company_info, "buyer_info": buyer_info, "bank_info": bank_info,
            "terms": DEFAULT_TERMS, "items": self.items,
        }
        app.local_db.save_document(self.doc_type, company_info["name"], number, formatted, confirmed, state, pdf_path, xlsx_path)
        app.local_db.upsert_profile_local("company", company_info["name"], dict(company_info, **bank_info))
        app.local_db.upsert_profile_local("buyer", buyer_info["name"], buyer_info)

        status = "confirmed (server number)" if confirmed else "saved offline - number will be confirmed on next sync"
        self._notify(f"{self.doc_type} {formatted} generated ({status}).\nPDF: {pdf_path}")
        app.run_sync()

    def _notify(self, text):
        dlg = MDDialog(text=text, buttons=[])
        close_btn = MDFlatButton(text="OK", on_release=lambda x: dlg.dismiss())
        dlg.buttons = [close_btn]
        dlg.open()
