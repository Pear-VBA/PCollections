Attribute VB_Name = "TestPCollections"
'@TestModule
'@Folder("Tests")
Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("ItemExists")
Private Sub TestItemExistsTrue()
    On Error GoTo TestFail
    
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add "1a"
    Coll.Add "2b"
    Coll.Add "3c"

    'Assert:
    Assert.IsTrue pcol_PCollections.pcol_ItemExists(Coll, "2b")

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ItemExists")
Private Sub TestItemExistsFalse()
    On Error GoTo TestFail
    
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add "1a"
    Coll.Add "2b"
    Coll.Add "3c"

    'Assert:
    Assert.IsFalse pcol_PCollections.pcol_ItemExists(Coll, "2a")

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("KeyExists")
Private Sub TestKeyExistsTrue()
    On Error GoTo TestFail
        
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add Item:=1, Key:="1a"
    Coll.Add Item:=2, Key:="2b"
    Coll.Add Item:=3, Key:="3c"

    'Assert:
    Assert.IsTrue pcol_PCollections.pcol_KeyExists(Coll, "2b")

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("KeyExists")
Private Sub TestKeyExistsFalse()                        'TODO Rename test
    On Error GoTo TestFail
    
    
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add "1a"
    Coll.Add "2b"
    Coll.Add "3c"

    'Assert:
    Assert.IsFalse pcol_PCollections.pcol_KeyExists(Coll, "2a")

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Join")
Private Sub TestJoin()
    On Error GoTo TestFail
    
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add "1a"
    Coll.Add "2b"
    Coll.Add "3c"

    'Assert:
    Assert.AreEqual "1a, 2b, 3c", pcol_PCollections.pcol_Join(Coll, ", ")

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Split")
Private Sub TestSplit()
    On Error GoTo TestFail
    
    Dim Coll As Collection
    Set Coll = pcol_PCollections.pcol_Split("1a, 2b, 3c", ", ")

    'Assert:
    Assert.AreEqual "1a", Coll(1)
    Assert.AreEqual "2b", Coll(2)
    Assert.AreEqual "3c", Coll(3)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ToArray")
Private Sub TestToArray()
    On Error GoTo TestFail
    
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add "1a"
    Coll.Add "2b"
    Coll.Add "3c"

    'Assert:
    Assert.SequenceEquals Array("1a", "2b", "3c"), pcol_PCollections.pcol_ToArray(Coll)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("FromArray")
Private Sub TestFromArray()
    On Error GoTo TestFail
    
    Dim Coll As Collection
    Set Coll = pcol_PCollections.pcol_FromIterable( _
        Array("1a", "2b", "3c") _
    )

    'Assert:
    Assert.AreEqual Conversion.CLng(3), Coll.Count
    Assert.AreEqual "1a", Coll(1)
    Assert.AreEqual "2b", Coll(2)
    Assert.AreEqual "3c", Coll(3)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("FromArray")
Private Sub TestFromEmptyArray()
    On Error GoTo TestFail
    
    Dim Coll As Collection
    Set Coll = pcol_PCollections.pcol_FromIterable(Array())

    'Assert:
    Assert.AreEqual Conversion.CLng(0), Coll.Count

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("FromArray")
Private Sub TestNonArray()
    Const ExpectedError As Long = 13
    On Error GoTo TestFail
    
    Dim Coll As Collection
    Set Coll = pcol_PCollections.pcol_FromIterable(4)

    'Assert:
    Dim Dummy As Long
    Dummy = Coll.Count
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Extend")
Private Sub TestExtend()
    On Error GoTo TestFail
    
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
    
    'Assert:
    Assert.AreEqual Conversion.CLng(6), Coll1.Count
    Assert.AreEqual "1a", Coll1(1)
    Assert.AreEqual "2b", Coll1(2)
    Assert.AreEqual "3c", Coll1(3)
    Assert.AreEqual "4d", Coll1(4)
    Assert.AreEqual "5e", Coll1(5)
    Assert.AreEqual "6f", Coll1(6)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Extend")
Private Sub TestExtendMissingDestination()
    Const ExpectedError As Long = 424
    On Error GoTo TestFail
    
    'Arrange:
    Dim Coll1 As Collection

    Dim Coll2 As Collection
    Set Coll2 = New Collection

    Coll2.Add "4d"
    Coll2.Add "5e"
    Coll2.Add "6f"

    pcol_PCollections.pcol_Extend Coll1, Coll2
    'Act:
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("FindAll")
Private Sub TestFindAll()
    On Error GoTo TestFail
    
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
    
    'Assert:
    Assert.AreEqual Conversion.CLng(2), Coll2.Count
    Assert.AreEqual "2b", Coll2(1)
    Assert.AreEqual "2b", Coll2(2)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("FindIndex")
Private Sub TestFindIndex()
    On Error GoTo TestFail
    
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add "1a"
    Coll.Add "2b"
    Coll.Add "3c"

    'Assert:
    Assert.AreEqual Conversion.CLng(2), pcol_PCollections.pcol_FindIndex(Coll, "2b")
    Assert.AreEqual Conversion.CLng(-1), pcol_PCollections.pcol_FindIndex(Coll, "4e")

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("FindLastIndex")
Private Sub TestFindLastIndex()
    On Error GoTo TestFail
    
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add "1a"
    Coll.Add "2b"
    Coll.Add "3c"
    Coll.Add "1a"
    Coll.Add "2b"
    Coll.Add "3c"

    'Assert:
    Assert.AreEqual Conversion.CLng(5), pcol_PCollections.pcol_FindLastIndex(Coll, "2b")
    Assert.AreEqual Conversion.CLng(-1), pcol_PCollections.pcol_FindLastIndex(Coll, "4e")

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Max")
Private Sub TestMax()
    On Error GoTo TestFail
    
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add 123
    Coll.Add 3
    Coll.Add 482
    Coll.Add 69
    Coll.Add 1
    Coll.Add 4

    'Assert:
    Assert.AreEqual 482, pcol_PCollections.pcol_Max(Coll)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Max")
Private Sub TestMaxStrings()
    On Error GoTo TestFail
    
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add "f"
    Coll.Add "a"
    Coll.Add "b"
    Coll.Add "c"

    'Assert:
    Assert.AreEqual "f", pcol_PCollections.pcol_Max(Coll)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Min")
Private Sub TestMin()
    On Error GoTo TestFail
    
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add 123
    Coll.Add 3
    Coll.Add 482
    Coll.Add 69
    Coll.Add 1
    Coll.Add 4

    'Assert:
    Assert.AreEqual 1, pcol_PCollections.pcol_Min(Coll)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Min")
Private Sub TestMinStrings()
    On Error GoTo TestFail
    
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add "f"
    Coll.Add "a"
    Coll.Add "b"
    Coll.Add "c"

    'Assert:
    Assert.AreEqual "a", pcol_PCollections.pcol_Min(Coll)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Pop")
Private Sub TestPop()
    On Error GoTo TestFail
    
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add 123
    Coll.Add 3
    Coll.Add 482
    Coll.Add 69
    Coll.Add 1
    Coll.Add 4

    'Assert:
    Assert.AreEqual 4, pcol_PCollections.pcol_Pop(Coll)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Pop")
Private Sub TestPopEmtyCollection()
    Const ExpectedError As Long = 9
    On Error GoTo TestFail
    
    Dim Coll As Collection
    Set Coll = New Collection

    'Assert:
    Dim Dummy As Variant
    Dummy = pcol_PCollections.pcol_Pop(Coll)
    
Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("Reverse")
Private Sub TestReverse()
    On Error GoTo TestFail
    
    Dim Coll1 As Collection
    Set Coll1 = New Collection

    Coll1.Add "1a"
    Coll1.Add "2b"
    Coll1.Add "3c"

    Dim Coll2 As Collection
    Set Coll2 = pcol_PCollections.pcol_Reverse(Coll1)

    'Assert:
    Assert.AreEqual Coll1(3), Coll2(1)
    Assert.AreEqual Coll1(2), Coll2(2)
    Assert.AreEqual Coll1(1), Coll2(3)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
