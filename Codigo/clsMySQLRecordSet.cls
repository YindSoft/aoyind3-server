VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMySQLRecordSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************
'clsMySQLRecordSet.cls - ImperiumAO - v1.4.5 - www.imperiumao.com.ar
'
'Interfaz RecordSet para MySQL
'
'*****************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with This program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

'*****************************************************************
'Augusto Rando(barrin@imperiumao.com.ar)
'   - First relase
'*****************************************************************

#If BattleServer = 0 Then

Option Explicit

'ADO Compat
Public Enum MySQLRecordsetStatus
    adStateOpen = 0
    adStateClosed = 1
End Enum

Private Type tDataType
    value As String
    Lenght As Long
End Type

Private lngRecords As Long
Private lngFields As Long

Private lngCurRow As Long
Private lngCurField As Long

Private arrData() As tDataType
Private arrColumns() As String

Private RecordsetStatus As MySQLRecordsetStatus

Public Sub Initialize(ByVal lngRecordsAffected As Long, ByVal lngNumFields As Long)

lngRecords = lngRecordsAffected
lngFields = lngNumFields
lngCurField = 0

ReDim arrData(lngFields - 1, lngRecords - 1)
ReDim arrColumns(lngFields - 1)

RecordsetStatus = adStateOpen

End Sub

Public Sub AddField(ByVal strField As String)

arrColumns(lngCurField) = Trim(UCase$(strField))
lngCurField = lngCurField + 1

End Sub

Public Sub MoveNext()

lngCurRow = lngCurRow + 1

End Sub

Public Function EOF() As Boolean

EOF = (lngCurRow > lngRecords - 1)

End Function

Public Sub SetValue(ByVal lngCol As Long, ByVal lngRow As Long, strValue As String, ByVal lngLenght As Long)

arrData(lngCol, lngRow).value = strValue
arrData(lngCol, lngRow).Lenght = lngLenght

End Sub

Public Property Get value(ByRef strField As String) As String
Attribute value.VB_UserMemId = 0

Dim lngColIndex As Long

If RecordsetStatus = adStateClosed Then Exit Property

lngColIndex = Col_Find(UCase$(strField))

If lngColIndex = -1 Then
    value = vbNullString
Else
    value = arrData(lngColIndex, lngCurRow).value
End If

End Property

Public Property Get valuebyindex(ByRef lngColIndex As Long) As String

If RecordsetStatus = adStateClosed Then Exit Property

If lngColIndex = -1 Then
    valuebyindex = vbNullString
Else
    valuebyindex = arrData(lngColIndex, lngCurRow).value
End If

End Property

Public Property Get State() As MySQLRecordsetStatus

State = RecordsetStatus

End Property

Public Sub CloseRecordset()

RecordsetStatus = adStateClosed

End Sub

Private Function Col_Find(ByVal strCol As String) As Long

On Error GoTo errhandler
    Dim i As Long
    
    Col_Find = -1
    
    For i = 0 To lngFields - 1
        If arrColumns(i) = strCol Then
            Col_Find = i
            Exit Function
        End If
    Next i
    
    Exit Function
    
errhandler:

End Function

Private Sub Class_Initialize()

RecordsetStatus = adStateClosed

End Sub

#End If
