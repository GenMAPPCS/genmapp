VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClusterGene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim id As String
Dim primaryIDs As New Collection
Dim masterPrimary As String
Dim GOterms As New Collection
Dim NestedGoterms As New Collection
Dim visited As Boolean
Dim prob As Double 'a number between 0 and 1 for picking R random genes

Public Sub setID(i As String)
   id = i
End Sub

Public Function getID() As String
   getID = id
End Function

Public Sub setMasterPrimary(id As String)
   masterPrimary = id
End Sub
Public Function getMasterPrimary() As String
   getMasterPrimary = masterPrimary
End Function
Public Sub addPrimaryID(primary As String)
   On Error GoTo error
   primaryIDs.Add primary, primary
error:
End Sub

Public Function getPrimaryIDs() As Collection
   Set getPrimaryIDs = primaryIDs
End Function
Public Sub addGOterm(GOID As String, term As goterm)
   On Error GoTo error
   GOterms.Add term, GOID
   NestedGoterms.Add term, GOID
error:
   'a duplicate GOID which is fine, but we just ignore it.
End Sub
Public Sub addLocalGOterm(GOID As String, term As goterm)
   On Error GoTo error
   GOterms.Add term, GOID
error:
   'a duplicate GOID which is fine, but we just ignore it.
End Sub
Public Function getGOTerms() As Collection
   Set getGOTerms = GOterms
End Function

Public Sub addNestedGOterm(GOID As String, term As goterm)
   On Error GoTo error
   NestedGoterms.Add term, GOID
error:
   
End Sub
Public Function hasGOID(GOID As String) As Boolean
On Error GoTo error
   Dim go As goterm
   
   Set go = NestedGoterms.Item(GOID)
   hasGOID = True
   Exit Function
error:
   'this GOID wasn't found
   hasGOID = False
End Function

Public Sub reset()
   visited = False
End Sub

Public Function getNestedGOterms() As Collection
   Set getNestedGOterms = NestedGoterms
End Function
'this class will store all of the ClusterSystem-GO relationships
'again we store this memory to avoid db tables
Public Function wasVisited() As Boolean
   wasVisited = visited
End Function

Public Sub visit()
   visited = True
End Sub
Private Sub Class_Initialize()
   visited = False
End Sub

Public Sub setProb(p As Double)
   prob = p
End Sub

Public Function getProb() As Double
   getProb = prob
End Function
