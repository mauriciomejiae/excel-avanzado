VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmRegistro 
   Caption         =   "UserForm1"
   ClientHeight    =   7155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7155
   OleObjectBlob   =   "FrmRegistro.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "FrmRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnGuardar_Click()

ActiveSheet.Cells(2, 1).Select
Selection.EntireRow.Insert

ActiveSheet.Cells(2, 1) = cmbTipoIdentificacion
ActiveSheet.Cells(2, 2) = txtNumeroIdentificacion.Value
ActiveSheet.Cells(2, 3) = txtNombre
ActiveSheet.Cells(2, 4) = txtApellidos
ActiveSheet.Cells(2, 5) = txtFechaNacimiento
ActiveSheet.Cells(2, 6) = txtTelefono.Value
ActiveSheet.Cells(2, 7) = txtDireccion
ActiveSheet.Cells(2, 8) = txtEmail

cmbTipoIdentificacion = Empty
txtNumeroIdentificacion = Empty
txtNombre = Empty
txtApellidos = Empty
txtFechaNacimiento = Empty
txtTelefono = Empty
txtDireccion = Empty
txtEmail = Empty

End Sub

Private Sub UserForm_Activate()

cmbTipoIdentificacion.AddItem "C.C"
cmbTipoIdentificacion.AddItem "T.I"
cmbTipoIdentificacion.AddItem "C.C"

txtNumeroIdentificacion.SetFocus

End Sub

Private Sub btnCerrar_Click()
End
End Sub

