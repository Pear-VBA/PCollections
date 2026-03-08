Attribute VB_Name = "Examples"
'@Folder "PCollectionsProject"
Option Explicit

Sub T()
    Dim Coll As Collection
    Set Coll = New Collection
    Coll.Add "A1"
    Coll.Add "A4"
    Coll.Add "A2", before:=2
    Coll.Add "A3", before:=3
    Debug.Print pcol_PCollections.pcol_Join(Coll, ", ")
End Sub

Public Sub ExampleItemExists()
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add "1a"
    Coll.Add "2b"
    Coll.Add "3c"

    Debug.Print pcol_PCollections.pcol_ItemExists(Coll, "2b") ' True
    Debug.Print pcol_PCollections.pcol_ItemExists(Coll, "2a") ' False
End Sub

Public Sub ExampleKeyExists()
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add Item:=1, Key:="1a"
    Coll.Add Item:=2, Key:="2b"
    Coll.Add Item:=3, Key:="3c"

    Debug.Print pcol_PCollections.pcol_KeyExists(Coll, "2b") ' True
    Debug.Print pcol_PCollections.pcol_KeyExists(Coll, "2a") ' False
End Sub

Public Sub ExampleJoin()
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add "1a"
    Coll.Add "2b"
    Coll.Add "3c"

    Debug.Print pcol_PCollections.pcol_Join(Coll, ", ") ' 1a, 2b, 3c
End Sub

Public Sub ExampleSplit()
    Dim Coll As Collection
    Set Coll = pcol_PCollections.pcol_Split("1a, 2b, 3c", ", ")

    Debug.Print pcol_PCollections.pcol_Join(Coll, ", ") ' 1a, 2b, 3c
End Sub

Public Sub ExampleToArray()
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add "1a"
    Coll.Add "2b"
    Coll.Add "3c"

    Debug.Print Strings.Join(PCollection.ToArray(Coll), ", ") ' 1a, 2b, 3c
End Sub

Public Sub ExampleFromIterable()
    Dim Coll As Collection
    Set Coll = pcol_PCollections.pcol_FromIterable( _
        Array("1a", "2b", "3c") _
    )

    Debug.Print Coll.Count ' 3
    Debug.Print pcol_PCollections.pcol_Join(Coll, ", ") ' 1a, 2b, 3c

    Dim Coll2 As Collection
    Set Coll2 = New Collection
    Coll2.Add "1a"
    Coll2.Add "2b"
    Coll2.Add "3c"
    Set Coll = pcol_PCollections.pcol_FromIterable(Coll2)

    Debug.Print Coll.Count ' 3
    Debug.Print pcol_PCollections.pcol_Join(Coll, ", ") ' 1a, 2b, 3c


    Range("A1").Value = "1a"
    Range("A2").Value = "2b"
    Range("A3").Value = "3c"
    Set Coll = pcol_PCollections.pcol_FromIterable(Range("A1:A3"))

    Debug.Print Coll.Count ' 3
    Debug.Print pcol_PCollections.pcol_Join(Coll, ", ") ' 1a, 2b, 3c
End Sub

Public Sub ExampleExtend()
    Dim Coll1 As Collection
    Set Coll1 = New Collection

    Coll1.Add "1a"
    Coll1.Add "2b"
    Coll1.Add "3c"

    Dim Coll2 As Collection
    Set Coll2 = New Collection

    Coll2.Add "4d"
    Coll2.Add "5e"
    Coll2.Add "6f"

    pcol_PCollections.pcol_Extend Coll1, Coll2
    Debug.Print pcol_PCollections.pcol_Join(Coll1, ", ") ' 1a, 2b, 3c, 4d, 5e, 6f
End Sub

Public Sub ExampleFindAll()
    Dim Coll1 As Collection
    Set Coll1 = New Collection

    Coll1.Add "1a"
    Coll1.Add "2b"
    Coll1.Add "3c"
    Coll1.Add "1a"
    Coll1.Add "2b"
    Coll1.Add "3c"

    Dim Coll2 As Collection
    Set Coll2 = pcol_PCollections.pcol_FindAll(Coll1, "2b")
    Debug.Print pcol_PCollections.pcol_Join(Coll2, ", ") ' 2b, 2b
End Sub

Public Sub ExampleFindIndex()
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add "1a"
    Coll.Add "2b"
    Coll.Add "3c"

    Debug.Print pcol_PCollections.pcol_FindIndex(Coll, "2b") ' 2
    Debug.Print pcol_PCollections.pcol_FindIndex(Coll, "4e") ' -1
End Sub

Public Sub ExampleFindLastIndex()
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add "1a"
    Coll.Add "2b"
    Coll.Add "3c"
    Coll.Add "1a"
    Coll.Add "2b"
    Coll.Add "3c"

    Debug.Print pcol_PCollections.pcol_FindLastIndex(Coll, "2b") ' 5
    Debug.Print pcol_PCollections.pcol_FindLastIndex(Coll, "4e") ' -1
End Sub

Public Sub ExampleMax()
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add 123
    Coll.Add 3
    Coll.Add 482
    Coll.Add 69
    Coll.Add 1
    Coll.Add 4

    Debug.Print pcol_PCollections.pcol_Max(Coll) ' 482
End Sub

Public Sub ExampleMin()
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add 123
    Coll.Add 3
    Coll.Add 482
    Coll.Add 69
    Coll.Add 1
    Coll.Add 4

    Debug.Print pcol_PCollections.pcol_Min(Coll) ' 1
End Sub

Public Sub ExamplePop()
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add 123
    Coll.Add 3
    Coll.Add 482
    Coll.Add 69
    Coll.Add 1
    Coll.Add 4

    Dim Item As Long
    Item = pcol_PCollections.pcol_Pop(Coll)
    Debug.Print Item ' 4
End Sub

Public Sub ExampleReverse()
    Dim Coll1 As Collection
    Set Coll1 = New Collection

    Coll1.Add "1a"
    Coll1.Add "2b"
    Coll1.Add "3c"
    Coll1.Add "4d"

    Dim Coll2 As Collection
    Set Coll2 = pcol_PCollections.pcol_Reverse(Coll1)
    Debug.Print pcol_PCollections.pcol_Join(Coll2, ", ") ' 3c, 2b, 1a
End Sub
