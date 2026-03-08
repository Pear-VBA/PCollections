Attribute VB_Name = "pcol_PCollections"
'@Folder "PCollectionsProject.src"
Option Explicit

'@Description "Adds all items from the Source collection to the Destination collection."
Public Sub pcol_Extend(ByRef Destination As Collection, ByRef Source As Collection)
    If Destination Is Nothing Then Information.Err().Raise 424
    If Source Is Nothing Then Exit Sub
    If Source.Count() = 0 Then Exit Sub

    Dim Item As Variant
    For Each Item In Source
        Destination.Add Item
    Next
End Sub

'@Description "Finds and returns a new collection containing all occurrences of the specified item in the Source collection."
Public Function pcol_FindAll(ByRef SourceCollection As Collection, ByVal Item As Variant) As Collection
    Dim Items As Collection: Set Items = New Collection

    Dim Check As Variant
    For Each Check In SourceCollection
        If Information.IsObject(Item) Then
            If Item Is Check Then Items.Add Check
        Else
            If Item = Check Then Items.Add Check
        End If
    Next

    Set pcol_FindAll = Items
End Function

'@Description "Returns the index of the first occurrence of the specified item in the Source collection, or -1 if not found."
Public Function pcol_FindIndex(ByRef SourceCollection As Collection, ByVal Item As Variant) As Long
    Dim i As Long
    For i = 1 To SourceCollection.Count()
        Dim Found As Boolean
        If Information.IsObject(SourceCollection(i)) Then
            Found = SourceCollection(i) Is Item
        Else
            Found = SourceCollection(i) = Item
        End If

        If Found Then
            pcol_FindIndex = i
            Exit Function
        End If
    Next

    pcol_FindIndex = -1
End Function

'@Description "Returns the index of the last occurrence of the specified item in the Source collection, or -1 if not found."
Public Function pcol_FindLastIndex(ByRef SourceCollection As Collection, ByVal Item As Variant) As Long
    Dim i As Long
    For i = SourceCollection.Count() To 1 Step -1
        Dim Found As Boolean
        If Information.IsObject(SourceCollection(i)) Then
            Found = SourceCollection(i) Is Item
        Else
            Found = SourceCollection(i) = Item
        End If

        If Found Then
            pcol_FindLastIndex = i
            Exit Function
        End If
    Next

    pcol_FindLastIndex = -1
End Function

'@Description "Creates and returns a new collection from the given array."
Public Function pcol_FromArray(ByRef SourceArray As Variant) As Collection
    Set pcol_FromArray = pcol_PCollections.pcol_FromIterable(SourceArray)
End Function

'@Description "Creates and returns a new collection from the given iterable."
Public Function pcol_FromIterable(ByRef Source As Variant) As Collection
    Dim Buffer As Collection: Set Buffer = New Collection
    Dim Item As Variant
    For Each Item In Source
        Buffer.Add Item
    Next

    Set pcol_FromIterable = Buffer
End Function

'@Description "Checks if the specified item exists in the Source collection and returns True or False."
Public Function pcol_ItemExists(ByRef SourceCollection As Collection, ByVal Item As Variant) As Boolean
    Dim Check As Variant
    For Each Check In SourceCollection
        Dim IsSame As Boolean
        If Information.IsObject(Item) Then
            IsSame = Item Is Check
        Else
            IsSame = Item = Check
        End If

        If IsSame Then
            pcol_ItemExists = True
            Exit Function
        End If
    Next
End Function

'@Description "Concatenates all items in the Source collection into a string, separated by the specified delimiter."
Public Function pcol_Join(ByRef SourceCollection As Collection, Optional ByVal Delimiter As String = " ") As String
    Dim Items As Variant: Items = pcol_PCollections.pcol_ToArray(SourceCollection:=SourceCollection)
    pcol_Join = Strings.Join(Items, Delimiter)
End Function

'@Description "Checks if a specified key exists in the Source collection and returns True or False."
Public Function pcol_KeyExists(ByRef SourceCollection As Collection, ByVal Key As Variant) As Boolean
    On Error Resume Next
    Dim Dummy As Boolean
    Dummy = Information.IsObject(SourceCollection(Key))

    pcol_KeyExists = Information.Err().Number = 0
End Function

'@Description "Returns the maximum value in the Source collection, raising an error if it contains objects."
Public Function pcol_Max(ByRef SourceCollection As Collection) As Variant
    Dim MaxItem As Variant
    Dim Item As Variant
    For Each Item In SourceCollection
        If Information.IsObject(Item) Then
            Information.Err().Raise _
                Number:=5, _
                Source:="pcol_PCollections.pcol_Max", _
                Description:="Function Max can't compare objects"
        End If

        If Information.IsEmpty(MaxItem) Then
            MaxItem = Item
        Else
            If MaxItem < Item Then MaxItem = Item
        End If
    Next

    pcol_Max = MaxItem
End Function

'@Description "Returns the minimum value in the Source collection, raising an error if it contains objects."
Public Function pcol_Min(ByRef SourceCollection As Collection) As Variant
    Dim MinItem As Variant
    Dim Item As Variant
    For Each Item In SourceCollection
        If Information.IsObject(Item) Then
            Information.Err().Raise _
                Number:=5, _
                Source:="pcol_PCollections.pcol_Min", _
                Description:="Function Min can't compare objects"
        End If

        If Information.IsEmpty(MinItem) Then
            MinItem = Item
        Else
            If MinItem > Item Then MinItem = Item
        End If
    Next

    pcol_Min = MinItem
End Function

'@Description "Removes and returns the last item from the Source collection."
Public Function pcol_Pop(ByRef SourceCollection As Collection) As Variant
    If SourceCollection Is Nothing Then Information.Err().Raise 424
    If SourceCollection.Count() = 0 Then Information.Err().Raise 9

    Dim LastIndex As Long: LastIndex = SourceCollection.Count()

    If Information.IsObject(SourceCollection.Item(LastIndex)) Then
        Set pcol_Pop = SourceCollection.Item(LastIndex)
    Else
        pcol_Pop = SourceCollection.Item(LastIndex)
    End If

    SourceCollection.Remove Index:=LastIndex
End Function

'@Description "Splits a string by the specified delimiter and returns a collection of the resulting substrings."
Public Function pcol_Split(ByVal Expression As String, Optional ByVal Delimiter As String = " ") As Collection
    Set pcol_Split = pcol_PCollections.pcol_FromIterable(Strings.Split(Expression, Delimiter))
End Function

'@Description "Returns a new collection with items from the Source collection in reverse order."
Public Function pcol_Reverse(ByRef SourceCollection As Collection) As Collection
    If SourceCollection Is Nothing Then Information.Err().Raise 424
    If SourceCollection.Count() < 2 Then
        Set pcol_Reverse = SourceCollection
        Exit Function
    End If

    Dim Reversed As Collection: Set Reversed = New Collection
    Reversed.Add SourceCollection(SourceCollection.Count)
    Reversed.Add SourceCollection(1)

    Dim l As Long
    Dim r As Long
    l = 2
    r = SourceCollection.Count - 1

    While l <= r
        If l <> r Then
            Reversed.Add SourceCollection(l), before:=r
        End If
        Reversed.Add SourceCollection(r), before:=l
        l = l + 1
        r = r - 1
    Wend

    Set pcol_Reverse = Reversed
End Function

'@Description "Converts the Source collection into an array and returns it."
Public Function pcol_ToArray(ByRef SourceCollection As Collection) As Variant
    If SourceCollection Is Nothing Then
        pcol_ToArray = Array()
        Exit Function
    End If

    If SourceCollection.Count() = 0 Then
        pcol_ToArray = Array()
        Exit Function
    End If

    Dim Items() As Variant: ReDim Items(0 To SourceCollection.Count() - 1)

    Dim i As Long
    For i = 1 To SourceCollection.Count()
        If Information.IsObject(SourceCollection(i)) Then
            Set Items(i - 1) = SourceCollection(i)
        Else
            Items(i - 1) = SourceCollection(i)
        End If
    Next

    pcol_ToArray = Items
End Function
