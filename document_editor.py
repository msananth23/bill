<ProfilesScreen>:
    MDBoxLayout:
        orientation: "vertical"

        MDTopAppBar:
            title: "Profiles"
            left_action_items: [["arrow-left", lambda x: root.go_back()]]

        ScrollView:
            MDBoxLayout:
                orientation: "vertical"
                adaptive_height: True
                size_hint_y: None
                height: self.minimum_height
                padding: "8dp"
                spacing: "8dp"

                MDLabel:
                    text: "Companies (synced across devices)"
                    font_style: "Subtitle2"
                    size_hint_y: None
                    height: "28dp"
                MDBoxLayout:
                    id: company_list
                    orientation: "vertical"
                    size_hint_y: None
                    height: self.minimum_height
                    adaptive_height: True

                MDLabel:
                    text: "Buyers (synced across devices)"
                    font_style: "Subtitle2"
                    size_hint_y: None
                    height: "28dp"
                MDBoxLayout:
                    id: buyer_list
                    orientation: "vertical"
                    size_hint_y: None
                    height: self.minimum_height
                    adaptive_height: True
