VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Processamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub Class_Initialize()
    'Mostra a Apulheta
    Screen.MousePointer = vbHourglass
    fFrmMDIPrincipal.stbBarraDeEstado.Panels(3).Text = "Processamento em curso, espere um momento S.F.F."
End Sub

Private Sub Class_Terminate()
    'Mostra o Normal
    Screen.MousePointer = vbDefault
    fFrmMDIPrincipal.stbBarraDeEstado.Panels(3).Text = vbNullString
End Sub
