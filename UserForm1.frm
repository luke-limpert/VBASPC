VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3876
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3720
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Dim lRow As Long

If UserForm1.Username.Value = "sbe" Then

        If UserForm1.Password.Value = "hook2" Then
        
         'Finds the last non-blank cell in a single row or column
        
             UserForm1.Hide
             UserForm2.Hide
    
         'Find the last non-blank cell in column A(1)
             lRow = Cells(Rows.Count, 1).End(xlUp).Row
    
             lRow = lRow + 1
    
             Range("A" & lRow).Select
    
             ActiveCell.FormulaR1C1 = "Stephanie Burnett"
        
             UserForm1.Username.Value = ""
        
             UserForm1.Password.Value = ""
            
        End If

UserForm1.Hide
    
    End If
    
If UserForm1.Username.Value = "Carmley" Then

        If UserForm1.Password.Value = "Rachel" Then
        
         'Finds the last non-blank cell in a single row or column
        
             UserForm1.Hide
             UserForm2.Hide
    
         'Find the last non-blank cell in column A(1)
             lRow = Cells(Rows.Count, 1).End(xlUp).Row
    
             lRow = lRow + 1
    
             Range("A" & lRow).Select
    
             ActiveCell.FormulaR1C1 = "Mitchell Carmley"
        
             UserForm1.Username.Value = ""
        
             UserForm1.Password.Value = ""
            
        End If

UserForm1.Hide
    
    End If
    
If UserForm1.Username.Value = "mohrr" Then

        If UserForm1.Password.Value = "1234" Then
        
         'Finds the last non-blank cell in a single row or column
        
             UserForm1.Hide
             UserForm2.Hide
    
         'Find the last non-blank cell in column A(1)
             lRow = Cells(Rows.Count, 1).End(xlUp).Row
    
             lRow = lRow + 1
    
             Range("A" & lRow).Select
    
             ActiveCell.FormulaR1C1 = "Ryan Mohr"
        
             UserForm1.Username.Value = ""
        
             UserForm1.Password.Value = ""
            
        End If
        
    End If
    
End Sub

