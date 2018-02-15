Module formSwitch
    
    Friend Enum FormName As Integer
        devForm = 0
        frmMTSend = 1 'Money Transfer
        frmPawning = 2
        frmInsurance = 3
        frmMTReceive = 4
        frmDollar = 5
        frmPawnItem = 6
        frmDollarSimple = 7
        frmMoneyExchange = 8
        frmAdminPanel = 9

        frmPawningV2_Client = 10
        frmPawningV2_Specs = 11
        frmPawningV2_Claimer = 12
        frmPawningV2_SpecsValue = 13
        frmPawningV2_InterestScheme = 14

        Coi = 15
        layAway = 16
        layAwayExist = 17

        frmWesternExchangeSender = 18
        frmWesternExchangeReceiver = 19
        frmWesternCurrency = 20

        POSClient = 21
    End Enum

    'Friend Sub ReloadFormFromSearch1(ByVal gotoForm As FormName, ByVal cr As Currency)
    '    Select Case gotoForm
    '        Case FormName.frmMoneyExchange
    '            frmmoneyexchange.LoadCurrencyall(cr)
    '        Case FormName.frmWesternCurrency
    '            frmWesternExchange.LoadCurrency(cr)
    '    End Select
    'End Sub
End Module