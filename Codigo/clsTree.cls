VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'*******************************************************************************************
'Pathfinding
'*******************************************************************************************
' Julien Lecomte
' webmaster@amanitamuscaria.org
' http://www.amanitamuscaria.org
' Feel free to use, abuse or distribute. (USUS, FRUCTUS, & ABUSUS)
' If you improve it, tell me !
' Don't take credit for what you didn't create. Thanks.
'*******************************************************************************************
'
' Vb Collection class & vb dictionnary class is *too* slow. To make things fast
' I'll use an array and add elements to the array.
'
' The A* method can be improved. It doesn't care whether a tile has a value of 1
' or 9, it just finds the fastest path, but not the shortest.
' For Dijkstra method, it finds the shortest path, but finds it pretty slowly.
' Anyhow, the point here isn't the speed of execution, but just the algo used
' I have found only one (complicated and slow) example of pathfinding under VB,
' so I guess this code is a good contribution to the VB community.
'
' If you want more speed under VB, cut out the graphic part of the code in frmPathfinding
' If you want even more speed, then shorten this class module and the code to either use only
' A* or Dijkstra.
' If you want even more speed, then don't use VB !
'
' Cheers,
' Julien
' 2001

'RVT  I've made a few changes in here - purely from a performance perspective
'     I think you'll find it quite a bit quicker at the expense of more storage (space-time continuum again)
'     The only code changes are in here, and in the new zinSort routine (©Me)  15 Jan 2002
'     Im sure it can still be made faster - but the fmrPathfinding will have to see too much rearrangement...
'     In particular the Path_GetValue routine is very expensive and is probably better put into this class
'     so that it can be made more efficient (esp now that a copy of AIMap is available in here). Some other
'     bits may also benefit because of the same.
'     Cheers,  Ron  15.1.2002

'RVT2 You wont get it faster than it is now I dont think. Sort is gone altogether as are nearly all of the
'     scanning lookups. Everything is pretty tight now.  55x faster for Dijkstra and 23x faster for A* than
'     when I started....
'     The path retrace is also now in the class so that there are now quite a few changes in frmPathFinding
'     Ive labelled things as usual
'     cheers Ron    17.1.2002

'RVT3 Even more found (in A* Algorithm) and now I have added complete algorithms (without display capability)
'     which are quickest of all. This is because the calling is reduced to a minimum. Ive folded the Explore
'     steps into Start and UpdateNode steps, and killed off a few other minor things. Changes in here and
'     in frmPathFinding. Variables have been renamed in here to make it a little more readable.
'     Ill leave you in peace now. Thanks for your patience

'     The alorithms are now O(N^2) now, where they were O(N^4).
'     by my estimation Dijkstra should be able to handle a 280x280 matrix in 1 sec when compiled in the worst case
'     for a maze-like path 320x320 should be possible in 1 sec  (on a 600MHz machine). That should be good enough
'     for a civil game... Other than very sparse matrices, Dijkstra will generally win.
'     21.1.2002
'
'
'*******************************************************************************************
'
' VERSION 2.0
'
' Dear All,
' I would like to thank Ron for his very positive feedback. His comments have been left in the code
' He contributed in a better sorting algorithm, heuristic and exploration of nodes. This deeply
' improved the performance of this vb algorithm.
'
' User Interface : You can't change settings, delete map, etc... anymore while pathfinding in progress.
'
' Any other changes are labeled "JL" in the code.
'
' Cheers,
' Julien
' webmaster@amanitamuscaria.org
'
'*******************************************************************************************
'
' VERSION 2.0 RC1
'
' Dear All,
' I'm releasing this code again, as it has become hard to find since several AI websites
' are now down.
' No code was changed, but the useless "draw tiles" section was removed.
'
' Julien
' 2004

Private Type tNode
  X          As Integer
  Y          As Integer
  lValue     As Long        'the Path value (sum) so far
  lHeuristic As Long        'a weighted value for choices
  lLink      As Long        'a link between subsequent nodes in the chain
End Type

Private Const LOW_BOUND As Long = 1

Public HIGH_BOUND      As Long

Private tNodes()       As tNode
Private lNodes         As Long    'The number of nodes in operation
Private lNodeCur       As Long    'The index of the one we are working on
Private lFirstEmpty    As Long    'Where to start searching for an empty node

Private lBeenHere()    As Long    '0 if not, if so index into tNodes(), -2 if impassable
Private lAIMap()       As Byte    'The connection set aiMAP

Private lEndX          As Integer
Private lEndY          As Integer    'To store the end points
Private lStartX          As Integer
Private lStartY          As Integer

Private xDim           As Long
Private yDim           As Long    'RVT The X and Y dimensions of the tableau (it doesnt need to be square anymore)

Dim xLD As Integer, xUD As Integer, yLD As Integer, yUD As Integer

Private Sub Class_Terminate()
    Erase tNodes
    Erase lBeenHere
    Erase lAIMap
End Sub

Public Function RunSearch(ByVal map As Byte, ByVal lXStart As Integer, ByVal lYStart As Integer, ByVal lXEnd As Integer, ByVal lYEnd As Integer, Optional ByVal AguaValida As Byte = 0) As Boolean
    '// Returns TRUE if path found
  
    StartSearch map, lXStart, lYStart, lXEnd, lYEnd, AguaValida
    
'AStar using the NextNode routine
        Do Until NextNode Or RunSearch
            RunSearch = UpdateCurrentNode
            DoEvents
        Loop
    BackTracePath                     'Go back from goal to start and record
End Function

'================================================================================================================

Public Sub StartSearch(map As Byte, lXStart As Integer, lYStart As Integer, lXEnd As Integer, lYEnd As Integer, Optional AguaValida As Byte = 0)
  
    Dim X As Integer
    Dim Y As Integer
    
    'Call FindLegalPosNpc(map, lXEnd, lYEnd, 0, True)
    
    
    If lXStart > lXEnd Then
        xUD = lXStart + 20
        xLD = lXEnd - 20
    Else
        xLD = lXStart - 20
        xUD = lXEnd + 20
    End If
    If lYStart > lYEnd Then
        yUD = lYStart + 20
        yLD = lYEnd - 20
    Else
        yLD = lYStart - 20
        yUD = lYEnd + 20
    End If
    
    
    
    If xLD < 5 Then xLD = 5
    If xUD > 1095 Then xUD = 1095
    
    xDim = xUD - xLD + 1
    
    If yLD < 5 Then yLD = 5
    If yUD > 1495 Then yUD = 1495
    
    yDim = yUD - yLD + 1
    
    'take a copy of the connection map - the costs of crossing a tile
    ReDim lAIMap(0 To xDim, 0 To yDim) As Byte
    For X = xLD To xUD
        For Y = yLD To yUD
            lAIMap(X - xLD, Y - yLD) = IIf(LegalPosNPC(map, X, Y, AguaValida, True) Or (X = lXStart And Y = lYStart) Or (X = lXEnd And Y = lYEnd), 1, 2)
        Next Y
    Next X
    
    'A record of which Nodes have been visited ie. are on the queue
    ReDim lBeenHere(0 To xDim, 0 To yDim) As Long
    
    HIGH_BOUND = xDim * yDim
    ReDim tNodes(LOW_BOUND To HIGH_BOUND) As tNode
  
    '// Store the end point
    lEndX = lXEnd - xLD
    lEndY = lYEnd - yLD
    lStartX = lXStart
    lStartY = lYStart
    
    '// Store the start point
    
    Add lXStart - xLD, lYStart - yLD 'This one now Current
    lNodeCur = LOW_BOUND 'Make it the current one
    With tNodes(lNodeCur)
        .lValue = 0          'StartofPath then Fist 4 nodes of near neigbours
        If .X > 0 Then Add .X - 1, .Y '// IMP
        If .Y > 0 Then Add .X, .Y - 1 '// IMP
        If (.X + 1) < xDim Then Add .X + 1, .Y
        If (.Y + 1) < yDim Then Add .X, .Y + 1
    End With
    
    lNodeCur = lNodeCur + 1 'the first node to be examined is this one
    lFirstEmpty = lNodeCur
End Sub
Private Function IsWalkable(ByVal map As Integer, ByVal row As Integer, ByVal Col As Integer) As Boolean
IsWalkable = MapData(map, row, Col).Blocked = 0 And Not HayAgua(map, row, Col) And MapData(map, row, Col).UserIndex = 0

End Function
Private Sub Add(lX As Integer, lY As Integer)
    'RVT First check if the point is in the search queue - if so forget it else flag it as in the array
    'RVT Ignore it if its going to be impassible anyway
    'RVT This really helps out the both algorithms
    'RVT we turn a linear time in N seek in the array into a constant time operation
    'RVT and limit the nodes checked to only those of use to the answer
    Dim lDistX As Long, lDistY As Long
    
    If lBeenHere(lX, lY) = 0 Then                'Is it already there ie. seen?
        If lAIMap(lX, lY) <> 2 Then  'and its a useful point to move to ie dont add duds
            'havent been here and it isn't impassable so add it
            lNodes = lNodes + 1
            With tNodes(lNodes)
            .X = lX
            .Y = lY
            .lValue = -1
            .lLink = 0                             'to keep track of where we came from
    
            '// When adding a point to the list, figure out a heuristic if in A* mode
                lDistX = lAIMap(lX, lY) * (lX - lEndX) 'increase the disadvantage of expensive paths
                lDistY = lAIMap(lX, lY) * (lY - lEndY)
                .lHeuristic = (lDistX * lDistX + lDistY * lDistY)
            End With
            lBeenHere(lX, lY) = lNodes       'mark as seen and in queue at this place
        Else
            lBeenHere(lX, lY) = -2     'seen but not in queue (impassable)
        End If
    End If
End Sub

Public Function NextNode() As Boolean
    Dim i As Long
        
        'RVT The A* algorithm looks for the EMPTY_PATH node with the smallest heuristic value.
        '    So if we keep track of the first empty one and then look for the
        '    smallest from there we save ourselves a sort and one and a bit linear time searches
        
        For i = lFirstEmpty To lNodes    'look for the first empty node from where we last looked
            If tNodes(i).lValue = -1 Then Exit For
        Next

        If i <= lNodes Then              'we found one
            lFirstEmpty = i               'This says where now to look from next time
            lNodeCur = SmallestEmptyPath 'This one is therefore the SMALLEST HEURISTIC EMPTY_PATH NODE
            Exit Function
        End If
    
    NextNode = True 'At this point the search is at an end - It may not have been successful (deadended)...
                    'BTW searches dont always end here - if the end point is encountered they stop anyway
End Function

Private Function SmallestEmptyPath() As Long 'Find the smallest Heuristic Path Value from lFirstEmpty on
    Dim i As Long, minHeuristic As Long, minNode As Long
    
    minNode = lFirstEmpty
    minHeuristic = tNodes(lFirstEmpty).lHeuristic       'remember where and what the Heuristic value is
    
    For i = lFirstEmpty + 1 To lNodes                   'now check the rest keeping track of the smallest lHeuristic
        If tNodes(i).lValue = -1 Then          'must be empty
            If tNodes(i).lHeuristic < minHeuristic Then 'and its smaller
                minNode = i
                minHeuristic = tNodes(i).lHeuristic
            End If
        End If
    Next
  
    If minNode <> lFirstEmpty Then
        SwapNodes minNode, lFirstEmpty
    End If                                            'we swap them around so that the pending list gradually shrinks
  
    SmallestEmptyPath = lFirstEmpty
End Function

Private Sub SwapNodes(ByVal lNodeIdx1 As Long, ByVal lNodeIdx2 As Long)
    Dim z As tNode
    
    'do the swap
    z = tNodes(lNodeIdx1)
    tNodes(lNodeIdx1) = tNodes(lNodeIdx2)
    tNodes(lNodeIdx2) = z
    
    'fix the BeenHere pointers
    lBeenHere(tNodes(lNodeIdx1).X, tNodes(lNodeIdx1).Y) = lNodeIdx1
    lBeenHere(tNodes(lNodeIdx2).X, tNodes(lNodeIdx2).Y) = lNodeIdx2
End Sub

Public Function BackTracePath()
On Error Resume Next
    Dim NodeFrom As Long, NodeTo As Long, z As Long
    
    'We here take the opportunity to straighten out the path for subsequent use
    'by moving backwards from the goal and reversing the linkages until we get to the start
    'This should be called after the path_search is over
    
    z = lBeenHere(lEndX, lEndY)
    If z Then lNodeCur = z
    NodeTo = lNodeCur
    NodeFrom = tNodes(lNodeCur).lLink
    tNodes(NodeTo).lLink = 0                 'mark as end of chain
    If NodeFrom Then
        Do
            z = tNodes(NodeFrom).lLink             'need this for a while
            tNodes(NodeFrom).lLink = NodeTo        'mark the new link
            NodeTo = NodeFrom
            NodeFrom = z
        Loop Until NodeFrom = 0                     'start has lLink=0
    End If
End Function

Public Function PathStepNext(ByRef lX As Integer, ByRef lY As Integer) As Long
On Error Resume Next
    lNodeCur = tNodes(lBeenHere(lX - xLD, lY - yLD)).lLink
    If lNodeCur Then GetNode lX, lY
    PathStepNext = lNodeCur
End Function

Public Function GetNode(ByRef lX As Integer, ByRef lY As Integer)
    lX = tNodes(lNodeCur).X + xLD
    lY = tNodes(lNodeCur).Y + yLD
End Function

Public Function GetNodeValue(ByVal lX As Integer, ByVal lY As Integer) As Long
    '// Cannot use Current Node
    Dim i As Long
                                            'lBeenHere() records the index of Node(x,y) in the nodelist
    i = lBeenHere(lX, lY)                   'a constant time lookup
    If i > 0 Then                           'its in the queue
        GetNodeValue = tNodes(i).lValue
    Else
        GetNodeValue = -1
    End If
End Function

'We update the shortest path to here - looking at neighbour nodes we already have
'When we can we expand by adding all of those neighbours that are not already there

Public Function UpdateCurrentNode() As Boolean      'TRUE if this node was the END Node
    
    Dim lVal As Long, lPathMinVal As Long
    Dim i As Long, X As Integer, Y As Integer
    
    lPathMinVal = 2147483247                  'assume a really large value
    
    With tNodes(lNodeCur)
    For X = .X - 1 To .X + 1 Step 2        'Check LEFT and RIGHT
        If X >= 0 Then
            If X < xDim Then
                i = lBeenHere(X, .Y)             'CHECK point
                If i = 0 Then
                    Add X, .Y                      'if not there Then Expand to it for next time
                ElseIf i > 0 Then
                    lVal = tNodes(i).lValue
                    If lVal >= 0 Then
                        If lVal < lPathMinVal Then
                            lPathMinVal = lVal
                            .lLink = i                'new link from here
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    For Y = .Y - 1 To .Y + 1 Step 2        'CHECK UP and DOWN
        If Y >= 0 Then
            If Y < yDim Then
                i = lBeenHere(.X, Y)             'CHECK point
                If i = 0 Then
                    Add .X, Y                      'if not there Then Expand to it for next time
                ElseIf i > 0 Then
                    lVal = tNodes(i).lValue
                    If lVal >= 0 Then
                        If lVal < lPathMinVal Then
                            lPathMinVal = lVal
                            .lLink = i                'new link from here
                        End If
                    End If
                End If
            End If
        End If
    Next
    .lValue = lPathMinVal + lAIMap(.X, .Y)                'currently min Path Cost
    UpdateCurrentNode = ((.X = lEndX) And (.Y = lEndY))   'TRUE if we found the END
    End With
End Function
