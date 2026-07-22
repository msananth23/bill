<SettingsScreen>:
    MDBoxLayout:
        orientation: "vertical"
        padding: "16dp"
        spacing: "12dp"

        MDTopAppBar:
            title: "Settings"
            left_action_items: [["arrow-left", lambda x: root.go_back()]]

        MDLabel:
            text: "Sync Server (your VPS)"
            font_style: "Subtitle2"
            size_hint_y: None
            height: "28dp"

        MDTextField:
            id: server_url
            hint_text: "https://your-vps-domain.com"
            helper_text: "Leave blank to run fully offline on this device"
            helper_text_mode: "persistent"

        MDTextField:
            id: api_key
            hint_text: "API Key"
            password: True
            helper_text: "Must match KSMS_API_KEY on the server"
            helper_text_mode: "persistent"

        MDRaisedButton:
            text: "Save"
            on_release: root.save()

        Widget:
