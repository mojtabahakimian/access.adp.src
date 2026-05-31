Version =131074
ColumnsShown =3
Begin
    Comment ="Attached to the Customer Labels Dialog form."
End
Begin
End
Begin
    Comment ="Attached to the PrintLabelsFor option group."
End
Begin
    MacroName ="Enable SelectCountry"
    Action ="RunCode"
    Comment ="If user select All Countries, do not enable the SelectCountry combo box."
    Argument ="Proc_Customer_Label_Enable_SelectCountr () "
End
Begin
End
Begin
    Comment ="Attached to the Preview button."
End
Begin
    MacroName ="Preview"
    Action ="RunCode"
    Comment ="Preview all records."
    Argument ="Proc_Customer_Labels_Dialog_Preview () "
End
Begin
    MacroName ="Print"
    Action ="RunCode"
    Comment ="Print all records."
    Argument ="Proc_Customer_Labels_Dialog_Print () "
End
Begin
    MacroName ="Cancel"
    Action ="Close"
    Comment ="Close the Customer Labels Dialog form."
    Argument ="2"
    Argument ="Customer Labels Dialog"
    Argument ="0"
End
