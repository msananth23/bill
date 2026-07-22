<DocumentEditorScreen>:
    ScrollView:
        MDBoxLayout:
            orientation: "vertical"
            padding: "16dp"
            spacing: "10dp"
            adaptive_height: True
            size_hint_y: None
            height: self.minimum_height

            MDTopAppBar:
                title: root.doc_type
                left_action_items: [["arrow-left", lambda x: root.go_back()]]

            MDLabel:
                text: "Company"
                font_style: "Subtitle2"
                size_hint_y: None
                height: "24dp"
            MDTextField:
                id: company_name
                hint_text: "Company Name"
            MDTextField:
                id: company_addr
                hint_text: "Company Address"
            MDTextField:
                id: company_gstin
                hint_text: "Company GSTIN"
            MDTextField:
                id: company_contact
                hint_text: "Company Contact"
            MDTextField:
                id: bank_name
                hint_text: "Bank Name"
            MDTextField:
                id: bank_ac
                hint_text: "Account No."
            MDTextField:
                id: bank_ifsc
                hint_text: "IFSC"

            MDLabel:
                text: "Buyer"
                font_style: "Subtitle2"
                size_hint_y: None
                height: "24dp"
            MDTextField:
                id: buyer_name
                hint_text: "Buyer Name"
            MDTextField:
                id: buyer_addr
                hint_text: "Buyer Address"
            MDTextField:
                id: buyer_gstin
                hint_text: "Buyer GSTIN (optional)"
            MDTextField:
                id: inv_date
                hint_text: "Document Date (DD-MM-YYYY)"

            MDLabel:
                text: "Add Item"
                font_style: "Subtitle2"
                size_hint_y: None
                height: "24dp"
            MDTextField:
                id: item_desc
                hint_text: "Description of Goods"
            MDBoxLayout:
                size_hint_y: None
                height: "56dp"
                spacing: "6dp"
                MDTextField:
                    id: item_hsn
                    hint_text: "HSN/SAC"
                MDTextField:
                    id: item_qty
                    hint_text: "Qty"
                    input_filter: "int"
            MDBoxLayout:
                size_hint_y: None
                height: "56dp"
                spacing: "6dp"
                MDTextField:
                    id: item_unit
                    hint_text: "Unit (Nos/Kgs/...)"
                    text: "Nos"
                MDTextField:
                    id: item_rate
                    hint_text: "Rate"
                    input_filter: "float"
                MDTextField:
                    id: item_gst
                    hint_text: "GST %"
                    text: "18"
                    input_filter: "float"
            MDRaisedButton:
                text: "+ Add Item"
                on_release: root.add_item()

            MDLabel:
                text: "Items (tap to remove)"
                font_style: "Subtitle2"
                size_hint_y: None
                height: "24dp"
            MDBoxLayout:
                id: item_list_box
                orientation: "vertical"
                size_hint_y: None
                height: self.minimum_height
                adaptive_height: True

            MDLabel:
                id: subtotal_label
                text: "Subtotal: Rs. 0.00"
                font_style: "Subtitle1"
                size_hint_y: None
                height: "28dp"

            MDRaisedButton:
                text: "Generate PDF + Excel"
                size_hint_x: 1
                md_bg_color: 0.2, 0.6, 0.3, 1
                on_release: root.generate_document()

            Widget:
                size_hint_y: None
                height: "24dp"
