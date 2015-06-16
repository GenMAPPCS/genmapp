VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GOterm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'if I knew VB better or was writing this in Java, I would make the tview.node object hold all of this.
'but since I can't do that, I'll store these objects and then get the data from them to label
'the tview.object.


Dim GOID As String
Dim genesInME As New Collection
Dim genesChangedInME As New Collection
Dim name As String
Dim visit As Boolean
Dim results(14) As Double
   '0 = changedlocal
   '1 = onchiplocal
   '2 = ingolocal
   '3 = percentlocal
   '4 = presentlocal
   '5 = changed
   '6 = onchip
   '7 = ingo
   '8 = percent
   '9 = present
   '10 = z score
   '11 = permute p
   '12 = adjusted p
   '13 = zerocount
Dim currentZ As Double 'used for p value calculation
Dim zeroZ As Double
Dim realZ As Double 'abs of Z observed in real data. Used for comparison against permutations
Dim permuteZ As Double
Dim monoZ As Double
Dim permutecount As Integer, permutep As Double
Dim monocount As Integer, monop As Double
Dim changed As Integer, zerocount As Integer

   
Public Sub setName(N As String)
   name = N
End Sub
Public Function getName() As String
   getName = name
End Function

Public Sub reset()
   results(0) = 0
   results(3) = 0
   results(5) = 0
   results(8) = 0
   results(10) = 0
   results(11) = 0
   results(12) = 0
   permutecount = 0
   monocount = 0
   realZ = 0
   monoZ = 0
   permuteZ = 0
   zeroZ = 0
   changed = 0
End Sub

Public Sub setGOID(id As String)
   GOID = id
End Sub

Public Function getGOID() As String
   getGOID = GOID
End Function



Public Sub addChangedGeen(id As String)
On Error GoTo error
   genesChangedInME.Add id, id
error:
End Sub
Public Function getChangedGenes() As Collection
   Set getChangedGenes = genesChangedInME
End Function
Public Sub setingolocal(num As Integer)
  results(2) = num
End Sub

Public Function getInGOLocal() As Integer
   getInGOLocal = results(2)
End Function

Public Sub setOnChipLocal(num As Integer)
   results(1) = num
End Sub

Public Function getOnChipLocal() As Integer
   getOnChipLocal = results(1)
End Function
 
Public Sub setChangedLocal(num As Integer)
   results(0) = num
End Sub

Public Function getChangedLocal() As Integer
   getChangedLocal = results(0)
End Function

Public Sub setingo(num As Integer)
   results(7) = num
End Sub

Public Function getInGO() As Integer
   getInGO = results(7)
End Function

Public Sub setOnChip(num As Integer)
   results(6) = num
End Sub

Public Function getOnChip() As Integer
   getOnChip = results(6)
End Function
 
Public Sub setChanged(num As Integer)
   results(5) = num
End Sub

Public Function getChanged() As Integer
   getChanged = results(5)
End Function

Private Sub Class_Initialize()
   results(0) = 0
   results(1) = 0
   results(2) = 0
   results(3) = 0
   results(4) = 0
   results(5) = 0
   results(6) = 0
   results(7) = 0
   results(8) = 0
   results(9) = 0
   results(10) = 0
   results(11) = 0
   results(12) = 0
   visit = False
   permutecount = 0
   monocount = 0
End Sub

Public Sub calculateLocal()
   If results(1) = 0 Then
      results(3) = 0
   Else
      results(3) = (results(0) / results(1)) * 100
   End If
   
   If results(2) = 0 Then
      results(4) = 0
   Else
      results(4) = (results(1) / results(2)) * 100
   End If
   
End Sub

Public Sub addGene(Gene As String)
   On Error GoTo error
   
   genesInME.Add Gene, Gene
   Exit Sub
error:
   'gene is duplicate. so don't add it.
End Sub

Public Sub addChangedGene(Gene As String)
   On Error GoTo errorchanged
   
   genesChangedInME.Add Gene, Gene
   Exit Sub
errorchanged:
   'gene is duplicate. so don't add it.
End Sub

Public Function getGenes() As Collection
   Set getGenes = genesInME
End Function



Public Function getGeneCount() As Integer
   getGeneCount = genesInME.count
End Function

Public Function getgeneschangedcount() As Integer
   getgeneschanged = genesChangedInME.count
End Function

Public Sub calculateResults()
   On Error GoTo error
   Debug.Print GOID
   results(8) = (results(5) / results(6)) * 100
   results(9) = (results(6) / results(7)) * 100

error:
   Select Case Err.Number
      Case 6 'overflow
         If GOID = "GO" Then
            MsgBox "Sorry, but MAPPFinder can not calculate any results for these data." _
               & " None of the genes in the dataset you loaded can be related to the Gene Ontology." _
               & " This means you either have a very small dataset of poorly annotated genes, or something" _
               & " is wrong with your current GenMAPP database. If you are only working with a subset of genes" _
               & " please load the entire dataset. If you have a large dataset (N > 1000) then contact GenMAPP.org" _
               & " and report this problem.", vbOKOnly
            End
         End If
   End Select

End Sub

Public Sub setZscore(z As Double)
   results(10) = z
   realZ = Abs(z)
End Sub

Public Function getResults() As Double()
   getResults = results
End Function

Public Function visited() As Boolean
   visited = visit
End Function

Public Function setvisited() As Boolean
   visit = True
End Function
'*********************************************************************************************************************
'The following functions are used for calculating P-values
'*********************************************************************************************************************
Public Sub CalculatePermuteZ(bigR As Long, bigN As Long)
   Dim numerator As Double
   Dim denom As Double
  ' Dim changed As Integer
   
   If GOID = "GO" Then
      permuteZ = 0
   Else
   
      If bigR - bigN = 0 Then
         'no z for 100% changed criteria
         permuteZ = 0
      ElseIf changed = 0 Then
         permuteZ = zeroZ
         zerocount = zerocount + 1
      Else
         'Debug.Print genes.count
         numerator = (changed - (results(6) * bigR / bigN))
         denom = Sqr(results(6) * (bigR / bigN) * (1 - (bigR / bigN)) * (1 - ((results(6) - 1) / (bigN - 1))))
         permuteZ = numerator / denom
      End If
   End If
   permuteZ = Abs(permuteZ) 'need to use abs because its two tailed
   If permuteZ >= realZ Then 'this permutation beat the real score, addone to permutecount for p value
      permutecount = permutecount + 1
   End If
End Sub
Public Function getPermuteZ() As Double
   getPermuteZ = permuteZ
End Function

Public Function getZscore() As Double
   getZscore = realZ 'return the abs(Z) for sorting in p value calculation
End Function





Public Sub resetME()
   changed = 0
   permuteZ = zeroZ
   'zerocount = 0
End Sub
Public Sub setZeroZ(bigR As Long, bigN As Long)
   'this sets the z score = to the value if 0 genes are changed in this term.
   'this is needed since mappfinder didn't calculate zs for terms with no genes
   'changed
   Dim numerator As Double
   Dim denom As Double
   Dim changed As Integer
   Dim currentZscore As Double
   
   
   If GOID = "GO" Then
      zeroZscore = 0
   Else
   
      If bigR - bigN = 0 Then
         'no zscore for 100% changed criteria
         zeroZscore = 0
      Else
         'Debug.Print genes.count
         numerator = (0 - (results(6) * bigR / bigN))
         denom = Sqr(results(6) * (bigR / bigN) * (1 - (bigR / bigN)) * (1 - ((results(6) - 1) / (bigN - 1))))
         currentZscore = numerator / denom
         zeroZ = currentZscore
         permuteZ = currentZscore
      End If
   End If
   
End Sub


Public Sub addOne()
   changed = changed + 1
End Sub

Public Sub setMonoZ(z As Double)
   monoZ = z
   If monoZ >= realZ Then
      monocount = monocount + 1
   End If
End Sub

Public Function getMonoZ() As Double
   getMonoZ = monoZ
End Function

Public Sub CalculatePValues(TRIALS As Long)
   permutep = permutecount / TRIALS
   results(11) = permutep
   monop = monocount / TRIALS
   
   results(12) = monop
End Sub

Public Sub setMonoP(p As Double)
   monop = p
   results(12) = monop
End Sub

Public Function getMonoP() As Double
   getMonoP = monop
End Function
