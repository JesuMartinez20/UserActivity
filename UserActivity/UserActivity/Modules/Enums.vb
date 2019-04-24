Module Enums
    'Hook Type'
    Public Enum HookType
        WH_KEYBOARD_LL = 13
        WH_MOUSE_LL = 14
    End Enum
    'Hook Codes'
    Public Enum HookCodes
        HC_ACTION = 0
    End Enum
    'Type of actions registered'
    Public Enum TypeAction
        ActivaApp = 1
        ScrollApp = 2
        CopyApp = 3
        PasteApp = 4
        TypeApp = 5
        'CombKeyApp = 6
    End Enum
End Module
