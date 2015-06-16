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

'9-08-03 - Nathan pointed out a problem of a single probe being counted multiple times in a
'GO term because that probe linked to multiple MGIs which were all in the same GO term.
'I was originally counting the number of probes meeting filter results[0] = results[0] + 1
'but now I'll rely on the genesLocal, GenesinMe, etc collections to count the distint number
'of probes/genes in this go term. This relies on the fact that a collection enforces that
'its keys are unique. -SD



Dim GOID As String
Dim genesInME As New Collection
Dim genesLocal As New Collection 'this is an unfortunate name, in local mapps geneslocal stores the cluter genes
'genes in the mapp. In the GO calculations, genesLocal = the genes specific to that go term, not the nested numbers.
Dim genesChangedInME As New Collection
Dim genesChangedLocal As New Collection
'Dim permutegenes As New Collection
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

Public Sub setValues()
   results(0) = genesChangedLocal.count
   results(1) = genesLocal.count
   results(5) = genesChangedInME.count
   results(6) = genesInME.count
End Sub
Public Sub setValuesLocal()
   results(0) = genesChangedLocal.count
   results(1) = genesLocal.count
   results(5) = genesChangedInME.count
   results(6) = genesLocal.count
   
End Sub
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
   Set genesChangedInME = New Collection
   Set genesChangedLocal = New Collection
 '  Set permutegenes = New Collection
End Sub

Public Sub setGOID(id As String)
   GOID = id
End Sub

Public Function getGOID() As String
   getGOID = GOID
End Function

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

Public Function addGene(gene As String) As Boolean
   On Error GoTo error
   
   genesInME.Add gene, gene
   addGene = True
   Exit Function
error:
   addGene = False
   'gene is duplicate. so don't add it.
End Function
'used in local mapps to keep track of cluster genes.
Public Function addClusterGene(gene As String) As Boolean
    On Error GoTo error
        
    genesLocal.Add gene, gene
    addClusterGene = True
    Exit Function
error:
    addClusterGene = False
End Function
Public Function addGeneLocal(gene As String) As Boolean
   On Error GoTo error
   
   genesInME.Add gene, gene
   genesLocal.Add gene, gene
   addGeneLocal = True
   Exit Function
error:
   'gene is duplicate. so don't add it.
   addGeneLocal = False
End Function

Public Sub addChangedGene(gene As String)
   On Error GoTo errorchanged
   
   genesChangedInME.Add gene, gene
   Exit Sub
errorchanged:
   'gene is duplicate. so don't add it.
End Sub
Public Sub addChangedLocal(gene As String)
   On Error GoTo errorlocal
   genesChangedLocal.Add gene, gene
errorlocal:
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
   'Debug.Print GOID
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
Public Sub unvisit()
   visit = False
End Sub
Public Sub addonelocal()
   changed = changed + 1
End Sub
'*********************************************************************************************************************
'The following functions are used for calculating P-values
'*********************************************************************************************************************
Public Sub CalculatePermuteZ(bigR As Long, bigN As Long)
   Dim numerator As Double
   Dim denom As Double
 
   If GOID = "GO" Then
      permuteZ = 0
   Else
   
      If bigR - bigN = 0 Then
         'no z for 100% changed criteria
         permuteZ = 0
      ElseIf changed = 0 Then
         permuteZ = zeroZ
         'zerocount = zerocount + 1
      ElseIf changed = bigR Then
         permuteZ = 1
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
   visit = False
'   Set permutegenes = New Collection
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
         If (denom = 0) Then 'we make this an arbitrarily large number. we'll never acutally hit zero since all of the genes have this go term.
            currentZscore = numerator / 0.1
         Else
            currentZscore = numerator / denom
         End If
         zeroZ = currentZscore
         permuteZ = currentZscore
      End If
   End If
   
End Sub


Public Function addOne() As Boolean
  If Not visit Then
      changed = changed + 1
      visit = True
      addOne = True
   Else
      addOne = False
   End If

End Function

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
