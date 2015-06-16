VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrimaryGene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim id As String
Dim Cluster As New Collection 'this will most of the time have one element, but could have a lot
Dim Iwasvisited As Boolean
Dim prob As Double

Public Sub setID(newID As String)
   id = newID
End Sub

Public Function getID() As String
   getID = id
End Function

Public Sub addClusterGene(cg As ClusterGene)
   On Error GoTo error
   Cluster.Add cg, cg.getID
error: 'ignore duplicate cluster genes for a given primary gene
End Sub

Public Function getClusterGenes() As Collection
   Set getClusterGenes = Cluster
End Function

'this class stores the primary-ClusterSystem relationship. By storing this in memory rather than just
'querying a db table I'll be able to calculat the results faster and be able to run multiple runs
'without ever using the database

Public Sub setProb(p As Double)
   prob = p
End Sub

Public Function getProb() As Double
   getProb = prob
End Function


