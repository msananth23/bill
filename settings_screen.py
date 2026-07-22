#:import MDRaisedButton kivymd.uix.button.MDRaisedButton

<DashboardScreen>:
    MDBoxLayout:
        orientation: "vertical"
        padding: "16dp"
        spacing: "12dp"

        MDTopAppBar:
            title: "KSMS Billing"
            elevation: 2
            right_action_items: [["cog", lambda x: root.open_settings()], ["sync", lambda x: root.refresh_status()]]

        MDLabel:
            text: root.sync_status_text
            theme_text_color: "Secondary"
            size_hint_y: None
            height: "24dp"
            font_style: "Caption"

        MDLabel:
            text: "Create New Document"
            font_style: "H6"
            size_hint_y: None
            height: "36dp"

        MDRaisedButton:
            text: "+ Tax Invoice"
            size_hint_x: 1
            on_release: root.new_document("Tax Invoice")

        MDRaisedButton:
            text: "+ Proforma Invoice"
            size_hint_x: 1
            on_release: root.new_document("Proforma Invoice")

        MDRaisedButton:
            text: "+ Quotation"
            size_hint_x: 1
            on_release: root.new_document("Quotation")

        MDBoxLayout:
            size_hint_y: None
            height: "1dp"
            md_bg_color: 0.85, 0.85, 0.85, 1

        MDFlatButton:
            text: "Company & Buyer Profiles"
            size_hint_x: 1
            on_release: root.open_profiles()

        Widget:
