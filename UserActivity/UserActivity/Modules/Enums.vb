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
        CopyApp = 1
        PasteApp = 2
        TypeApp = 3
        ScrollApp = 4
        'CombKeyApp = 5
        ActivaApp = 6
    End Enum
End Module
