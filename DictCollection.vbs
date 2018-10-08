' Run with "cscript.exe DictCollection25.vbs"
' Comment out the next line to skip the SelfTest and Demo:
set dc = new DictCollection: dc.SelfTest(true): dc.DemoBasicFunctionality

' ===================================================================================================================================================
' DictCollection 026 (VBScript Version) - A mix between Dictionary and Collection that can emulate both
' Karsten Held 2016-2018, karstenheld3@gmail.com, MIT License
'
' ----------- HOW TO USE ----------------------------------------------------------------------------------------------------------------------------
' Creating an empty collection:                 Set dc = New DictCollection
' Adding an item with a key:                    dc.Add "key1", 123.45                           -> dc.Item(0) = 123.45
' Adding an item without a key:                 dc.Add, "A"                                     -> dc.Item(1) = "A"
' Getting the number or items:                  n = dc.Count                                    -> n = 2
' Getting the number or keys:                   k = dc.KeyCount                                 -> k = 1
' Getting an item by key:                       a = dc.Item("key1")                             -> a = 123.45
' Getting an item by index:                     a = dc.Item(0)                                  -> a = 123.45
'                                               b = dc.Item(1)                                  -> b = "A"
' Setting an item by index:                     dc.Item(0) = 100                                -> dc.Item(0) = 100
' Setting an item by key:                       dc.Item("key1") = 200                           -> dc.Item(0) = 200
' Adding an item by key:                        dc.Item("key2") = "B"                           -> dc.Item(2) = "B"
' Nesting another DictCollection:               dc.Add "key3", new DictCollection               -> dc.Item(3) = DictCollection
' Adding to nested DictCollection:              dc.Item("key3").Add "key4", 456.78              -> dc.Item(3).Item(0) = 456.78
'                                               dc.Item("key3").Add, "C"                        -> dc.Item(3).Item(1) = "C"
' Getting from nested DictCollection:           c = dc.Item("key3").Item("key4")                -> c = 456.78
'                                               d = dc.Item("key3").Item(1)                     -> d = "C"
' Checking if nested item exists:               dc.Item("key3").Item(1) = dc.NonExistingValue   -> false
'                                               dc.Item("key3").Item(2) = dc.NonExistingValue   -> true
' Remove an item by key:                        dc.Remove("key1")                               -> "A" is now at index 0
' Remove an item by index:                      dc.Remove(0)                                    -> "B" is now at index 0
' Getting an item's key (Empty if none):        a = dc.KeyOfItemAt(0)                           -> a = "key2"
' Checking if key exists:                       check = dc.Exists("key3")                       -> check = True
' Changing an items key:                        dc.Key("key2") = "key4"                         -> "key2" changed to "key4"
' Changing an items key while removing
'   another item that used this key before:     dc.Key("key4") = "key3"                         -> "key4" -> "key3", dc.Item(1) removed
' Getting keys for each item incl. Empty:       ka = dc.Keys                                    -> ka = Array("key3")
' Getting all items in the added order:         ia = dc.Items                                   -> ia = Array("B")
' Inserting an item with key at index=0:        dc.Insert "D", 0, "key5"                        -> dc.Item(0) = "D", dc.Item(1) = "B"
' Inserting an item without key at index=3:     dc.Insert "E", 3, ""                            -> dc.Item(2) = Empty, dc.Item(3) = "E"
' Getting all keys in sorted order:             ka = dc.SortedKeys                              -> ka = Array("key3", "key5")
' Removing all items from DictCollection:       dc.RemoveAll
'
' ----------- GLOBAL SETTINGS ------------------------------------------------------------------------------------------------------------------------
' dc.ThowErrors = False (default)                   will return .NonExistingValue when accessing nonexistent indexes or keys
' dc.ThowErrors = True                              will throw errors when accessing nonexistent indexes or keys
' dc.LazySorting = True (default)                   will sort key array at first access using a key and thus speed up .Add
' dc.LazySorting = False                            will sort key array at every .Add if key is used
' dc.CompareMode = 0 (default)                      keys are case-sensitive, will binary compare keys, fastest compare method
' dc.CompareMode = 1                                keys are case-insensitive, Ä = ä = A = a = á = Á
' dc.CompareMode = 2                                keys are case-insensitive (only MSAccess), uses localeID for matching
' dc.EmulateDictionary = True                       DictCollection will behave exactly like Scripting.Dictionary object
' dc.EmulateCollection = True                       DictCollection will behave exactly like VB Collection object
' dc.ZeroBasedIndex = True (default)                first item can be accessed with index = 0 (false -> first index = 1)
' dc.DefaultValueEnabled = False                    default property .Item() returns a reference to the DictCollection,
'                                                   allowing it to be used like a regular object: 'Set x = dc' and 'dc Is DictCollection' will work
' dc.DefaultValueEnabled = True (default)           default property .Item() returns "[EMPTY]" when DictCollection has no items and "[NONEXISTING]"
'                                                   when accessing nonexisting keys or indexes. This allows chaining of nonexisting item access.
' dc.NonExistingValue = "[NONEXISTING]" (default)   the value to be returned for nonexisting items if ThrowError=False
'                                                   can be used to overwrite this value with anything except objects, e.g. Empty or 0
' dc.EmptyCollectionValue = "[EMPTY]" (default)     the default value of empty DictCollections if ThrowError=False and DefaultValueEnabled=True
'                                                   can be used to overwrite this value with anything except objects, e.g. Empty or 0
' ===================================================================================================================================================
' KNOWN BUGS: None
' KNOWN PROBLEMS:
' - 'Is DictCollection' and 'VarType(dc)' cannot be used because default value evaluation returns String. Use TypeName() or '.EnableDefaultValue=False'.
' TODOS:
' - implement Property Let KeyOfItemAt(Index, Value)

Class FindKeyResult: Public Exact: Public Index: End Class
Class EvaluateIndexOrKeyResult: Public WasIndex: Public WasKey: Public Index: End Class
Class EvaluateKeyResult: Public WasNumber: Public WasString: Public Key: End Class

Class DictCollection

Private THROW_ERRORS_DEFAULT
Private COMPAREMODE_DEFAULT
Private  LAZYSORTING_DEFAULT
Private DEFAULTVALUEENABLED_DEFAULT
Private  ZEROBASEDINDEX_DEFAULT

' returned when ThrowErrors=False and acccessing a nonexisting key or index with default value: Set d = new DictCollection: wscript.echo d(0) -> "[NONEXISTING]"
Private  NONEXISTING_VALUE_DEFAULT
' returned when ThrowErrors=False and DefaultValueEnabled=True and acccessing an empty collection with default value: Set d = new DictCollection: wscript.echo d -> "[EMPTY]"
Private  EMPTY_COLLECTION_VALUE_DEFAULT

' "all" prefixes to put variables on top of property list in watch window when debugging
Private allItems
Private allItemKeyIndexes
Private allKeys
Private allKeyItemIndexes
Private mThrowErrors ' true: will throw errors when accessing nonexisting items or using invalid types
Private mCompareMode ' switches between case sensitive and case insensitive key comparison
' vbBinaryCompare = 0   default case sensitive comparison, sort order is derived from internal binary representation,
'                       unix-filename-like sorting (uppercase chars are sorted before lowercase chars)
' vbTextCompare = 1     case insensitive comparison, based on the order in the ASCII table
'                       will ignore all local character derivations: á = A = a = ä = Ä
' vbDatabaseCompare = 2 case insensitive, can only be used in Microsoft Access VBA
'                       will determine the sort order using the local ID of the database
Private mLazySorting  ' true: will not sort keys while adding but later when keys are searched
Private mKeysSorted  ' true: allKeys array is sorted
Private mExternalIndexOffset  ' 1: Collection compatible, 0: Scripting.Dictionary compatible
Private mEmulateCollection  ' if true, emulates Collection behavior
Private mEmulateScriptingDictionary  ' if true, emulates Scripting.Dictionary behavior
Private mDefaultValueEnabled ' if true, returns the first item as default value or if empty "[EMPTY]" and if nonexisting "[NONEXISTING"]
Private mNonExistingValue  ' returned when ThrowErrors=False and acccessing a nonexisting key or index with default value: Set d = new DictCollection: WscripT.echo d(0)
Private mEmptyCollectionValue  ' returned when ThrowErrors=False and acccessing an empty collection with default value: Set d = new DictCollection: WscripT.echo d
Private mItemCount  ' number of items. if item array is uninitialized (nonexisting state), then number of items = -1
Private mKeyCount  ' number of keys. if keys array is uninitialized (array state), then number of keys = -1
Private mScriptingDictionaryObjectKeys ' used by Scripting.Dictionary compatibility: holds object references where objects are used as keys
Private mScriptingDictionaryStoredKeyTypes ' used by Scripting.Dictionary compatibility: holds the key types of number keys (Integer, Long, Single, Double, Currency, Decimal, Date)


'VBScript Change: since VBScript does not support IsMissing we have to implement the function
Function IsMissing(p): IsMissing = (VarType(p) = vbError): End Function

' ======================== START: INITIALIZATION AND STATES =========================================================================================
Private Sub Class_Initialize(): init: End Sub
Private Sub Class_Terminate(): End Sub
Private Sub init()

    THROW_ERRORS_DEFAULT  = False
    LAZYSORTING_DEFAULT = True
    COMPAREMODE_DEFAULT = 0 ' Compatible with Scripting.Dictionary
    ZEROBASEDINDEX_DEFAULT  = True
    NONEXISTING_VALUE_DEFAULT = "[NONEXISTING]"
    DEFAULTVALUEENABLED_DEFAULT = True
    EMPTY_COLLECTION_VALUE_DEFAULT = "[EMPTY]"

    applyDefaultAccessSettings
    mLazySorting = LAZYSORTING_DEFAULT
    mKeysSorted = False
    initializeItemArray
    uninitializeKeyArray
End Sub
Private Sub applyDefaultAccessSettings()
    mThrowErrors = THROW_ERRORS_DEFAULT
    mCompareMode = COMPAREMODE_DEFAULT
    mNonExistingValue = NONEXISTING_VALUE_DEFAULT
    mEmptyCollectionValue = EMPTY_COLLECTION_VALUE_DEFAULT
    If ZEROBASEDINDEX_DEFAULT Then mExternalIndexOffset = 0 Else mExternalIndexOffset = 1
    mDefaultValueEnabled = DEFAULTVALUEENABLED_DEFAULT
End Sub
Private Sub initializeItemArray()
    allItems = Array()
    allItemKeyIndexes = Array()
    Erase allItemKeyIndexes
    mItemCount = 0
End Sub
Private Sub initializeKeyArray()
    allKeys = Split("", ".")
    allKeyItemIndexes = Array()
    Erase allKeyItemIndexes
    mKeyCount = 0
End Sub
Private Sub initializeDictionaryObjectKeyArray(): mScriptingDictionaryObjectKeys = Array(): End Sub
Private Sub initializeDictionaryStoredKeyTypes(): Set mScriptingDictionaryStoredKeyTypes = New DictCollection: mScriptingDictionaryStoredKeyTypes.NonExistingValue = "": End Sub
Private Sub uninitializeKeyArray()
    allKeys = Array(): allKeyItemIndexes = Array()
    Erase allKeys
    Erase allKeyItemIndexes
    mKeyCount = -1
End Sub
Private Sub uninitializeItemArray()
    allItems = Array() : allItemKeyIndexes = Array() 
    Erase allItems
    Erase allItemKeyIndexes
    mItemCount = -1
End Sub
Private Sub uninitializeDictionaryObjectKeyArray(): mScriptingDictionaryObjectKeys = Empty: End Sub
Private Sub uninitializeDictionaryStoredKeyTypes(): Set mScriptingDictionaryStoredKeyTypes = Nothing: End Sub
' 0 = nonexisting, 1 = empty array, 2 = filled array, 3 = empty key/value store, 4 = filled key/value store, 5 = filled key/value store with some items having no key
Public Property Get CollectionType()
    If mItemCount < 0 Then
        CollectionType = 0 ' nonexisting
    ElseIf mItemCount = 0 Then
        If mKeyCount < 0 Then
            CollectionType = 1 ' empty array
        Else
            CollectionType = 3 ' empty key/value store
        End If
    Else
        If mKeyCount < 0 Then
            CollectionType = 2 ' filled array
        Else
            Dim i
            For i = 0 To (mItemCount - 1)
                If allItemKeyIndexes(i) = -1 Then
                    CollectionType = 5 ' filled key/value store with some items having no key
                    Exit Property
                End If
            Next
            CollectionType = 4 ' filled key/value store
        End If
    End If
End Property
Public Property Let CollectionType(Value)
    Select Case Value
        Case 0 ' nonexisting
            uninitializeItemArray
            uninitializeKeyArray
        Case 1 ' empty array
            initializeItemArray
            uninitializeKeyArray
        Case 2 ' filled array
            If mItemCount < 0 Then initializeItemArray
            If mKeyCount > 0 Then setAllItemKeyIndexesToNone
            If mKeyCount > -1 Then uninitializeKeyArray
        Case 3 ' empty key/value store
            initializeItemArray
            initializeKeyArray
        Case 4 ' filled key/value store
            If mItemCount < 0 Then initializeItemArray
            If mKeyCount < 0 Then initializeKeyArray
            EnsureAllItemsHaveKeys
        Case 5 ' filled key/value store with some items having no key
            If mItemCount < 0 Then initializeItemArray
            If mKeyCount < 0 Then initializeKeyArray
    End Select
End Property
' sets all item-to-key pointer to -1 (none), needed to remove all keys
Private Sub setAllItemKeyIndexesToNone()
    Dim i
    For i = 0 To mItemCount: allItemKeyIndexes(i) = -1: Next
End Sub
' sets the keys of all items that have no key to "_IDX" where IDX = the items index, e.g. key for item at 12 = "_12"
' if "_IDX" is already taken, a number N ranging from 0 to 1000 is added as "_IDX_N" until the new key is not taken
Public Sub EnsureAllItemsHaveKeys()
    Dim i, keyToTest, keySubCounter, result
    ' sort if necessary
    If mLazySorting And mKeyCount > 1 And (Not mKeysSorted) Then sortKeysAndRemoveDuplicates: mKeysSorted = True
    ' initialize key array if necessary
    If mKeyCount = -1 Then initializeKeyArray
    For i = 0 To mItemCount
        If allItemKeyIndexes(i) = -1 Then
            ' test key
            keyToTest = "_" & CStr(i)
            Set result = findKeyIndex(keyToTest)
            keySubCounter = 0
            While result.Exact And keySubCounter <= 1000
                ' test key with subnumber
                keyToTest = "_" & i & "_" & CStr(keySubCounter)
                Set result = findKeyIndex(keyToTest)
                keySubCounter = keySubCounter + 1
            Wend
            If (Not result.Exact) And (keySubCounter <= 1000) Then
                ' add key using eager sorting
                insertKey keyToTest, result.Index, i
            Else
                If mThrowErrors Then Err.Raise 5, "DictCollection", "Could not assign key to item at " & i & ". Tried keys from '" & "_" & i & "' to '" & keyToTest & "'.": Exit Sub
            End If
        End If
    Next
End Sub
Private Function increaseItemArray() 
    increaseItemArray = UBound(allItems) + 1
    ReDim Preserve allItems(increaseItemArray)
    ReDim Preserve allItemKeyIndexes(increaseItemArray)
    mItemCount = increaseItemArray + 1
End Function
Private Function increaseItemArrayTo(newUpperBound)
    Dim oldUpperBound, Index 
    oldUpperBound = UBound(allItems)
    If newUpperBound > oldUpperBound Then
        ReDim Preserve allItems(newUpperBound)
        ReDim Preserve allItemKeyIndexes(newUpperBound)
        For Index = oldUpperBound + 1 To newUpperBound
            allItemKeyIndexes(Index) = -1 ' set all new key indexes to -1
        Next
        mItemCount = newUpperBound + 1
    End If
End Function
Private Function decreaseItemArray() 
    decreaseItemArray = UBound(allItems) - 1
    If decreaseItemArray < 0 Then
        initializeItemArray
    Else
        ReDim Preserve allItems(decreaseItemArray)
        ReDim Preserve allItemKeyIndexes(decreaseItemArray)
        mItemCount = decreaseItemArray + 1
    End If
End Function
Private Function increaseKeyArray() 
    increaseKeyArray = UBound(allKeys) + 1
    ReDim Preserve allKeys(increaseKeyArray)
    ReDim Preserve allKeyItemIndexes(increaseKeyArray)
    mKeyCount = increaseKeyArray + 1
End Function
Private Function decreaseKeyArray() 
    decreaseKeyArray = UBound(allKeys) - 1
    If decreaseKeyArray < 0 Then
        uninitializeKeyArray
    Else
        ReDim Preserve allKeys(decreaseKeyArray)
        ReDim Preserve allKeyItemIndexes(decreaseKeyArray)
        mKeyCount = decreaseKeyArray + 1
    End If
End Function
Public Property Get IsNonExisting() : IsNonExisting = (mItemCount = -1): End Property
Public Property Let IsNonExisting(ByVal Value)
    If Value Then ' Set IsNonExisting=True
        If mItemCount > -1 Then
            uninitializeItemArray
            uninitializeKeyArray
        End If
    Else ' Set IsNonExisting=False
        If mItemCount = -1 Then init ' re-initialize to existing state
    End If
End Property
Public Property Get EmptyCollectionValue(): EmptyCollectionValue = mEmptyCollectionValue: End Property
Public Property Let EmptyCollectionValue(Value): mEmptyCollectionValue = Value: End Property
Public Property Get NonExistingValue(): NonExistingValue = mNonExistingValue: End Property
Public Property Let NonExistingValue(Value): mNonExistingValue = Value: End Property
Public Property Get CompareMode() : CompareMode = mCompareMode: End Property
Public Property Let CompareMode(Value)
    If mCompareMode <> Value Then
        If mKeyCount > 0 And mThrowErrors Then
            ' Emulate Scripting.Dictionary behavior where change of CompareMode is only possible when Dictionary is empty
            Err.Raise 5, "Invalid procedure call or argument": Exit Property
        Else
            mCompareMode = Value
            If mKeyCount > 0 Then sortKeys 0, (mKeyCount - 1) 're-sort keys so that findKeyIndex will work as intended
        End If
    End If
End Property
Public Property Get ThrowErrors() : ThrowErrors = mThrowErrors: End Property
Public Property Let ThrowErrors(Value): mThrowErrors = Value: End Property
Public Property Get LazySorting() : LazySorting = mLazySorting: End Property
Public Property Let LazySorting(Value)
    mLazySorting = Value
    ' sort if necessary
    If Not mLazySorting And mKeyCount > 1 And (Not mKeysSorted) Then sortKeysAndRemoveDuplicates: mKeysSorted = True
End Property
Public Property Get ZeroBasedIndex() : ZeroBasedIndex = (mExternalIndexOffset = 0): End Property
Public Property Let ZeroBasedIndex(Value)
    if Value=True then
        mExternalIndexOffset = 0
    else
        mExternalIndexOffset = 1
    end if
End Property
Public Property Get DefaultValueEnabled(): DefaultValueEnabled = mDefaultValueEnabled: End Property
Public Property Let DefaultValueEnabled(Value): mDefaultValueEnabled = Value: End Property
Public Property Get EmulateCollection() :  EmulateCollection = mEmulateCollection: End Property
Public Property Let EmulateCollection(Value)
    mEmulateCollection = Value
    If mEmulateCollection Then
        mCompareMode = vbTextCompare
        mExternalIndexOffset = 1
        mThrowErrors = True
        mEmulateScriptingDictionary = False
        uninitializeDictionaryObjectKeyArray
        uninitializeDictionaryStoredKeyTypes
    Else
        applyDefaultAccessSettings
        uninitializeDictionaryObjectKeyArray
        uninitializeDictionaryStoredKeyTypes
    End If
End Property
Public Property Get EmulateDictionary() :  EmulateDictionary = mEmulateScriptingDictionary: End Property
Public Property Let EmulateDictionary(Value)
    mEmulateScriptingDictionary = Value
    If mEmulateScriptingDictionary Then
        mExternalIndexOffset = 0
        mThrowErrors = True
        mNonExistingValue = Empty
        mEmptyCollectionValue = Empty
        mEmulateCollection = False
        initializeDictionaryObjectKeyArray
        initializeDictionaryStoredKeyTypes
    Else
        applyDefaultAccessSettings
        uninitializeDictionaryObjectKeyArray
        uninitializeDictionaryStoredKeyTypes
    End If
End Property

' copies all settings from one DictCollection to another that allow nested access with modified NonExistingValue and EmptyCollectionValue to work properly
Public Sub CopyDefaultValueSettings(fromCollection, toCollection)
    If (fromCollection Is Nothing) Or (toCollection Is Nothing) Then Exit Sub
    toCollection.ThrowErrors = fromCollection.ThrowErrors
    toCollection.NonExistingValue = fromCollection.NonExistingValue
    toCollection.EmptyCollectionValue = fromCollection.EmptyCollectionValue
End Sub

' copies all settings from one DictCollection to another
Public Sub CopyAllSettings(fromCollection, toCollection)
    If (fromCollection Is Nothing) Or (toCollection Is Nothing) Then Exit Sub
    toCollection.ThrowErrors = fromCollection.ThrowErrors
    toCollection.NonExistingValue = fromCollection.NonExistingValue
    toCollection.EmptyCollectionValue = fromCollection.EmptyCollectionValue
    toCollection.ZeroBasedIndex = fromCollection.ZeroBasedIndex
    toCollection.CollectionType = fromCollection.CollectionType
    toCollection.CompareMode = fromCollection.CompareMode
    toCollection.LazySorting = fromCollection.LazySorting
    toCollection.DefaultValueEnabled = fromCollection.DefaultValueEnabled
End Sub
' ======================== END: INITIALIZATION AND STATES ===========================================================================================


' VBScript Change: A separate default property with zero parameters is needed to allow chaining of nonexisting item access
Public Default Property Get DefaultProperty()
    If mDefaultValueEnabled Then
        Select Case mItemCount
            Case -1: DefaultProperty = mNonExistingValue ' uninitialized item array = (mItemCount=-1) = non existing
            Case 0: DefaultProperty = mEmptyCollectionValue ' 
            Case Else:
                if IsObject(allItems(0)) then set DefaultProperty = allItems(0) Else DefaultProperty = allItems(0) 
        End Select
    Else
        Set DefaultProperty = Me ' set to object if default value is disabled
    End If
End Property


' ======================== START: ITEM, ITEMAT, ADD, ADD2, EXISTS, REMOVE, REMOVEALL ================================================================
' Get item by index or key. If ThrowErrors = False (default), then
' 1) nonexisting indexes/keys will return placeholder string "[NONEXISTING]"
'    this allows chaining of nonexisting indexes/keys:
'    example: dc.Item("existing").Item("existing").Item("nonexisting").Item("nonexisting") will result in "[NONEXISTING]"
' 2) empty nested dictcollections will return placeholder string "[EMPTY]"
'    example: dc.Item("existing").Item("empty") will result in "[EMPTY]"
Public Property Get Item(IndexOrKey)
    Dim result, action ' 0 = found, 1 = nonexisting/not found, 2 = default value
    set result = new EvaluateIndexOrKeyResult
    If IsMissing(IndexOrKey) Then
        ' Emulate Collection behavior
        If mThrowErrors And Not mEmulateScriptingDictionary Then Err.Raise 5, "DictCollection", "Invalid procedure call or argument": Exit Property
        action = 2 ' default value
    ElseIf mItemCount > 0 Then
        action = 1 ' nonexisting/not found
        If mKeyCount > 1 And mLazySorting And (Not mKeysSorted) Then sortKeysAndRemoveDuplicates: mKeysSorted = True
        If mEmulateScriptingDictionary Then
            Set result = evaluateIndexOrKey(convertScriptingDictionaryToDictCollectionKey(IndexOrKey))
        ElseIf mEmulateCollection Then
            If TypeName(IndexOrKey) = "Nothing" Then
                Err.Raise 13, "DictCollection", "Type mismatch": Exit Property
            Else
                Set result = evaluateIndexOrKey(convertCollectionToDictCollectionKey(IndexOrKey))
            End If
        Else
            Set result = evaluateIndexOrKey(IndexOrKey)
        End If
        If result.Index < 0 Or result.Index >= mItemCount Then
            If result.WasIndex Then
                If mThrowErrors Then
                    ' Emulate Collection behavior
                    If mItemCount < 1 Then
                        Err.Raise 5, "DictCollection", "Invalid procedure call or argument": Exit Property
                    Else
                        Err.Raise 9, "DictCollection", "Subscript out of range": Exit Property
                    End If
                End If
                action = 1 ' nonexisting/not found
            ElseIf result.WasKey Then
                ' Key was not found. Scripting.Dictionary returns data type = "Empty" (zero length string)
                ' We return an unassigned return value which is always Empty by default
                action = 1 ' nonexisting/not found
            Else
                ' was neither string nor number
                If mThrowErrors Then
                    ' Emulate Collection behavior
                    If mItemCount < 1 Then
                        Err.Raise 5, "DictCollection", "Invalid procedure call or argument": Exit Property
                    Else
                        If IsEmpty(IndexOrKey) Then
                            Err.Raise 9, "DictCollection", "Subscript out of range": Exit Property
                        Else
                            Err.Raise 13, "DictCollection", "Type mismatch": Exit Property
                        End If
                    End If
                End If
                action = 1 ' nonexisting/not found
            End If
        Else
            action = 0 ' found
        End If
    Else
        action = 1 ' nonexisting/not found
    End If
    Select Case action
        Case 0: ' found
            If IsObject(allItems(result.Index)) Then Set Item = allItems(result.Index) Else Item = allItems(result.Index)
        Case 1: ' nonexisting/not found
            If mThrowErrors And Not mEmulateScriptingDictionary Then Err.Raise 5, "DictCollection", "Invalid procedure call or argument": Exit Property
            ' Emulate Scripting.Dictionary behavior of implicitly adding nonexisting keys with empty items when they are accessed
            If mEmulateScriptingDictionary Then Me.Add IndexOrKey, Empty
            If mDefaultValueEnabled Then
                Dim newDictCollection: Set newDictCollection = New DictCollection
                CopyDefaultValueSettings Me, newDictCollection
                newDictCollection.CollectionType = 0 ' set to nonexisting
                newDictCollection.ThrowErrors = False
                Set Item = newDictCollection
            Else
                Set Item = mNonExistingValue
            End If
        Case 2: ' default value
            If mDefaultValueEnabled Then
                If mItemCount = -1 Then Item = mNonExistingValue: Exit Property  ' nonexisting value
                If mItemCount = 0 Then
                    Item = mEmptyCollectionValue ' value of empty DictCollection
                Else
                    If IsObject(allItems(0)) Then Set Item = allItems(0) Else Item = allItems(0)
                End If
            Else
                Set Item = Me: Exit Property ' set to object if default value is disabled
            End If
    End Select
End Property

' Alternative (and fast) property for explicitly accessing an item by index (for instance when emulating Scripting.Dictionary and index numbers are interpreted as keys)
Public Property Get ItemAt(Index)
    ItemAt = mNonExistingValue
    If Index < 0 Or Index >= mItemCount Then
        ' Emulate Collection behavior
        If mThrowErrors Then Err.Raise 9, "DictCollection", "Subscript out of range": Exit Property
    Else
        If IsObject(allItems(Index)) Then Set ItemAt = allItems(Index) Else ItemAt = allItems(Index)
    End If
End Property

' Sets an item to basic type. If ThrowErrors=True and item does not exist, throws errors.
' If ThrowErrors=False (Default), will add the item by given index or key
Public Property Let Item(IndexOrKey, Value)
    If mEmulateScriptingDictionary Then
        Dim key2: key2 = convertScriptingDictionaryToDictCollectionKey(IndexOrKey)
        internalSetItem key2, Value
        storeScriptingDictionaryKeyInformation IndexOrKey, key2
    ElseIf mEmulateCollection Then
        internalSetItem convertCollectionToDictCollectionKey(IndexOrKey), Value
    Else
        internalSetItem IndexOrKey, Value
    End If
End Property

' Sets an item to object. If ThrowErrors=True and item does not exist, throws errors.
' If ThrowErrors=False (Default), will add the item by given index or key
Public Property Set Item(IndexOrKey, Value)
    If mEmulateScriptingDictionary Then
        Dim key2: key2 = convertScriptingDictionaryToDictCollectionKey(IndexOrKey)
        internalSetItem key2, Value
        storeScriptingDictionaryKeyInformation IndexOrKey, key2
    ElseIf mEmulateCollection Then
        internalSetItem convertCollectionToDictCollectionKey(IndexOrKey), Value
    Else
        internalSetItem IndexOrKey, Value
    End If
End Property

' Sets or adds an item for a given index or key
' This function is emulation independent (with the exception of throwing compatible errors)
Private Sub internalSetItem(IndexOrKey, Value)
    Dim result, action  ' 0 = found, 1 = not found/byIndex, 2 = not found/byKey, 3 = not found/default value
    set result = new EvaluateIndexOrKeyResult
    If IsMissing(IndexOrKey) Then
        If mItemCount > 0 Then
            ' set default item (first item)
            result.Index = 0: result.WasIndex = True
            action = 0 ' found
        Else
            action = 3 ' not found/default value
        End If
    Else
        If mKeyCount > 1 And mLazySorting And (Not mKeysSorted) Then sortKeysAndRemoveDuplicates: mKeysSorted = True
        set result = evaluateIndexOrKey(IndexOrKey)
        If result.Index < 0 Or result.Index >= mItemCount Then
            ' key or index was not found
            If result.WasIndex Then
                ' index was found
                If mThrowErrors Then
                    ' Emulate Collection/Array behavior
                    Err.Raise 9, "DictCollection", "Subscript out of range"
                    Exit Sub
                End If                
                action = 1 ' not found/byIndex
            ElseIf result.WasKey Then
                ' key was not found
                If mThrowErrors Then
                    ' Emulate Scripting.Dictionary behavior
                    Err.Raise 91, "DictCollection", "Object variable or With block variable not set"
                    Exit Sub
                End If
                action = 2 ' not found/byKey
            Else
                ' was neither string nor numbe
                If mThrowErrors Then
                    ' Emulate Scripting.Dictionary behavior
                    Err.Raise 91, "DictCollection", "Object variable or With block variable not set"
                End If
                ' fail silently
                Exit Sub
            End If
        End If
    End If
    Select Case action
        Case 0: ' found
            If IsObject(Value) Then
                Set allItems(result.Index) = Value
                If TypeName(Value) = "DictCollection" Then ' expression "Value Is DictCollection" cannot be used because of default value evaluation
                    Dim dc2
                    Set dc2 = Value
                    CopyDefaultValueSettings Me, dc2
                End If
            Else
                allItems(result.Index) = Value
            End If
        Case 1: ' not found/byIndex
            internalAddItem Value, "", CLng(IndexOrKey)
        Case 2: ' not found/byKey
            internalAddItem Value, CStr(IndexOrKey), -1
        Case 3: ' not found/default value
            internalAddItem Value, "", 0
    End Select
End Sub

' Compatible with Dictionary .Add function
Public Sub Add(Key, Item)
    Dim result: Set result = New EvaluateKeyResult
    If mEmulateScriptingdictionary Then
        Dim objectKeyIndex 
        result.Key = convertScriptingDictionaryToDictCollectionKey(Key)
        internalAddItem Item, result.Key, -1
        ' check if key was object and add it if it does not exist or re-set it, if it does exist
        objectKeyIndex = getScriptingDictionaryObjectKeyIndex(result.Key)
        If UBound(mScriptingDictionaryObjectKeys) > -1 Or objectKeyIndex > -1 Then
            If objectKeyIndex > UBound(mScriptingDictionaryObjectKeys) Then UtilAddArrayValue mScriptingDictionaryObjectKeys, Key
       End If
        ' check if key is number type and if yes, remember original key type name
        If isScriptingDictionaryStoredKeyType(result.Key) Then mScriptingDictionaryStoredKeyTypes.Item(result.Key) = TypeName(Key)
        Exit Sub
    ElseIf mEmulateCollection Then
        set result = evaluateKey(convertCollectionToDictCollectionKey(Key))
    Else
        If IsMissing(Key) Then Key=""
        set result = evaluateKey(Key)
    End If
    If (Not result.WasNumber) And (Not result.WasString) Then
        If mThrowErrors Then Err.Raise 13, "DictCollection", "Keys of type " & TypeName(Key) & " are not supported"
        Exit Sub
    End If
    internalAddItem Item, result.Key, GetMissingValue(, 1) 
    If mEmulateScriptingDictionary Then storeScriptingDictionaryKeyInformation Key, result.Key
End Sub

' Inserts an item at a given index
Public Sub Insert(Item, AtIndex, Key)
    if IsMissing(Item) then Item = Empty
    if IsMissing(AtIndex) then AtIndex = -1
    If mEmulateScriptingdictionary Then
        Dim key2, objectKeyIndex : key2 = convertScriptingDictionaryToDictCollectionKey(Key)
        internalAddItem Item, key2, AtIndex
        ' check if key was object and add it if it does not exist or re-set it, if it does exist
        objectKeyIndex = getScriptingDictionaryObjectKeyIndex(key2)
        If UBound(mScriptingDictionaryObjectKeys) > -1 Or objectKeyIndex > -1 Then
            If objectKeyIndex > UBound(mScriptingDictionaryObjectKeys) Then UtilAddArrayValue mScriptingDictionaryObjectKeys, Key
       End If
        ' check if key is number type and if yes, remember original key type name
        If isScriptingDictionaryStoredKeyType(key2) Then mScriptingDictionaryStoredKeyTypes.Item(key2) = TypeName(Key)
    Else
        if IsMissing(Key) then Key = ""
        internalAddItem Item, CStr(Key), AtIndex
    End If
    If mEmulateScriptingDictionary Then storeScriptingDictionaryKeyInformation Key, result.Key    
End Sub

' Compatible with Collection .Add function with the exception that the last parameter 'After' cannot be left unassigned (VBScript limitation) and has to be set explicitly to Null
' Example 1 (VBA):          dc.Add2 "MyItem"
' Example 1 (VBScript):     dc.Add2 "MyItem", , , Null
' Example 2 (VBA):          dc.Add2 "MyItem", "key1", 5
' Example 2 (VBScript):     dc.Add2 "MyItem", "key1", 5, Null
' Example 3 (VBA):          dc.Add2 "MyItem", After:=3
' Example 3 (VBScript):     dc.Add2 "MyItem", , , 3
Public Sub Add2(Item, Key, Before, After)
    Dim AtIndex, result, tempIndex, tname
    set result = new EvaluateKeyResult
    AtIndex = -1 ' undefined
    if IsMissing(Before) then
    else
        ' Emulate Collection error behavior
        If IsObject(Before) Then tname = "Object" Else tname = TypeName(Before)
        Select Case tname
            Case "String": If mThrowErrors Then Err.Raise 5, "DictCollection", "Invalid procedure call or argument": Exit Sub
            Case "Empty": If mThrowErrors Then Err.Raise 9, "DictCollection", "Subscript out of range": Exit Sub
            Case "Integer", "Long", "Single", "Double", "Currency", "Decimal", "Date":
                ' VBScript Change: -1 is special value for 'missing'
                if not IsNull(Before) then 
                    tempIndex = Before - mExternalIndexOffset
                    If mThrowErrors And (tempIndex < 0 Or tempIndex > (mItemCount - 1)) Then Err.Raise 9, "DictCollection", "Subscript out of range": Exit Sub
                    AtIndex = Before ' offset will be substracted later in internalSetItem
                end if
            Case "Null": 'do nothing
            Case Else:
                If mThrowErrors And Not IsMissing(Before) Then Err.Raise 13, "DictCollection", "Type mismatch": Exit Sub
        End Select
    end if
    if IsMissing(After) then
    else
        If IsObject(After) Then tname = "Object" Else tname = TypeName(After)
        Select Case tname
            Case "String": If mThrowErrors Then Err.Raise 5, "DictCollection", "Invalid procedure call or argument": Exit Sub
            Case "Empty": If mThrowErrors Then Err.Raise 9, "DictCollection", "Subscript out of range": Exit Sub
            Case "Integer", "Long", "Single", "Double", "Currency", "Decimal", "Date":
                ' VBScript Change: -1 is special value for 'missing'
                if AtIndex = -1 and not IsNull(After) then 
                    ' Emulate Collection behavior
                    tempIndex = After - mExternalIndexOffset + 1
                    If mThrowErrors And (tempIndex <= 0 Or tempIndex > mItemCount) Then Err.Raise 9, "DictCollection", "Subscript out of range": Exit Sub
                    AtIndex = After + 1 ' offset will be substracted later in internalSetItem
                end if
            Case "Null": 'do nothing
            Case Else:
                If mThrowErrors And Not IsMissing(After) Then Err.Raise 13, "DictCollection", "Type mismatch": Exit Sub
        End Select
    end if
    If mThrowErrors Then
        ' Emulate Collection error behavior
        If IsNumeric(Key) Or IsDate(Key) Then
            Dim result2: Set result2 = New EvaluateIndexOrKeyResult: Set result2 = evaluateIndexOrKey(Key)
            If mItemCount > 0 And (result2.Index >= 0 Or result2.Index <= (mItemCount - 1)) Then
                Err.Raise 457, "DictCollection", "This key is already associated with an element of this collection": Exit Sub
            Else
                Err.Raise 13, "DictCollection", "Type mismatch": Exit Sub
            End If
        Else
            If (TypeName(Key) <> "String") And (Not IsMissing(Key)) Then Err.Raise 13, "DictCollection", "Type mismatch": Exit Sub
        End If
    End If    
    If mEmulateCollection Then
        set result = evaluateKey(convertCollectionToDictCollectionKey(Key))
    Else
        set result = evaluateKey(Key)
    End If
    internalAddItem Item, result.Key, AtIndex
End Sub

' Adds item with given key ("" = no key) at given index (if an item exists at index, it will be pushed one up)
' This function is emulation independent (with the exception of throwing compatible errors)
Private Sub internalAddItem(Item ,  Key, AtIndex)
    if IsMissing(Item) then Item = Empty
    if IsMissing(Key) then Key=""
    if IsMissing(AtIndex) then AtIndex = -1
    Dim newKey, newItemIndex, lastItemIndex, newUpperBound, keyPassed, indexPassed
    set newKey = new FindKeyResult 
    keyPassed = Len(Key) > 0 Or mEmulateScriptingdictionary
    If keyPassed Then
        If mKeyCount = -1 Then initializeKeyArray 'convert array to collection if key is being used
        If mLazySorting And Not mThrowErrors Then
            If AtIndex > -1 Then
                ' key was passed and item will be put somewhere in item array -> sort before because key collisions can't be fixed after this insert
                If Not mKeysSorted And mKeyCount > 1 Then sortKeysAndRemoveDuplicates: mKeysSorted = True
                Set newKey = findKeyIndex(Key) 'find key
            Else
                ' put key to position after last
                newKey.Index = mKeyCount
                newKey.Exact = False
                mKeysSorted = False
            End If
        ElseIf mLazySorting And mThrowErrors Then
            ' key was passed and we need to check if key already exists -> sort before because otherwise we can't throw an error
            If Not mKeysSorted And mKeyCount > 1 Then sortKeysAndRemoveDuplicates: mKeysSorted = True
            Set newKey = findKeyIndex(Key) 'find key
        Else
            Set newKey = findKeyIndex(Key) 'find key
        End If
    Else
        newKey.Exact = False
    End If
    If newKey.Exact Then
        ' we only get here if key was passed and key already exists
        If mThrowErrors Then
            'throw error because of Scripting.Dictionary emulation
            Err.Raise 457, "DictCollection", "This key is already associated with an element of this collection"
            Exit Sub
        Else
            ' key exists, replace item
            If IsObject(Item) Then
                Set allItems(allKeyItemIndexes(newKey.Index)) = Item
            Else
                allItems(allKeyItemIndexes(newKey.Index)) = Item
            End If
        End If
    Else
        ' STEP 1: insert item in itemArray
        AtIndex = AtIndex - mExternalIndexOffset
        If mItemCount < 0 Then initializeItemArray
        lastItemIndex = mItemCount - 1
        If AtIndex > -1 Then
            ' index was passed
            If AtIndex > lastItemIndex + 1 Then
                ' expand item array to given new index
                increaseItemArrayTo AtIndex
            Else
                increaseItemArray
            End If
            newItemIndex = AtIndex
        Else
            newItemIndex = increaseItemArray
        End If
        If newItemIndex <= lastItemIndex Then
            ' copy items one up
            copyItemsUp newItemIndex, (mItemCount - 2), 1, (mKeyCount > 0)
        End If
        If IsObject(Item) Then
            Set allItems(newItemIndex) = Item
            ' if item is a DictCollection, copy default value settings
            If TypeName(Item) = "DictCollection" Then ' expression "Item Is DictCollection" cannot be used because of default value evaluation
                Dim dc2: Set dc2 = Item
                CopyDefaultValueSettings Me, dc2
            End If
        Else
            allItems(newItemIndex) = Item
        End If
        If keyPassed Then
            allItemKeyIndexes(newItemIndex) = newKey.Index
            ' STEP 2: insert key in keyArray
            insertKey Key, newKey.Index, newItemIndex
            ' invalidate sorting
            If mLazySorting Then mKeysSorted = False
        Else
            allItemKeyIndexes(newItemIndex) = -1 ' set to -1 = no key exists for this item
        End If
    End If
End Sub

' inserts key at given index after copying all existing keys >= keyIndex one up
Private Function insertKey(Key, keyIndex, itemIndex)
    increaseKeyArray
    copyKeysUp keyIndex, mKeyCount - 2, 1, True
    allKeys(keyIndex) = Key
    allKeyItemIndexes(keyIndex) = itemIndex
End Function

' Removes an item at a given index or key
Public Function Remove(IndexOrKey)
    If mItemCount < 0 Then Exit Function
    Dim result, found, key2
    Set result = new EvaluateIndexOrKeyResult 
    If mKeyCount > 1 And mLazySorting And (Not mKeysSorted) Then sortKeysAndRemoveDuplicates: mKeysSorted = True
    If mEmulateScriptingdictionary Then
        key2 = convertScriptingDictionaryToDictCollectionKey(IndexOrKey)
        set result = evaluateIndexOrKey(key2)
    ElseIf mEmulateCollection Then
        If TypeName(IndexOrKey) = "String" Then
            set result = evaluateIndexOrKey(convertCollectionToDictCollectionKey(IndexOrKey))
        Else
            set result = evaluateIndexOrKey(IndexOrKey)
        End If
    Else
        set result = evaluateIndexOrKey(IndexOrKey)
    End If
    If result.Index < 0 Or result.Index > (mItemCount - 1) Then
        If result.WasIndex Then
            ' Emulate Collection behavior
            If mThrowErrors Then
                If mItemCount < 1 Then
                    Err.Raise 5, "DictCollection", "Invalid procedure call or argument"
                Else
                    Err.Raise 9, "DictCollection", "Subscript out of range"
                End If
            End If
            Exit Function
        ElseIf result.WasKey Then
            ' key was not found
            If mThrowErrors And mEmulateCollection Then Err.Raise 5, "DictCollection", "Invalid procedure call or argument"
            ' Scripting.Dictionary and Tim Hall's Dictionary both throw this error
            If mThrowErrors Then Err.Raise 32811, "DictCollection", "Application-defined or object-defined error"
            Exit Function
        Else
            ' was neither string nor number
            If mThrowErrors Then
                ' Emulate Collection behavior
                If mItemCount < 1 Or IsMissing(IndexOrKey) Then
                    Err.Raise 5, "DictCollection", "Invalid procedure call or argument"
                Else
                    Err.Raise 9, "DictCollection", "Subscript out of range"
                End If
            End If
            Exit Function
        End If
    End If
    Dim keyIndex, hasKeys 
    hasKeys = mKeyCount > 0
    keyIndex = allItemKeyIndexes(result.Index)
    If (result.Index = (mItemCount - 1)) Then
        'last item, do nothing
    Else
        ' copy all items after found index one down
        copyItemsDown result.Index + 1, mItemCount - 1, 1, hasKeys
    End If
    decreaseItemArray
    If hasKeys And keyIndex > -1 Then
        If keyIndex = (mKeyCount - 1) Then
            ' last item, do nothing
        Else
            ' copy keys one down
            copyKeysDown keyIndex + 1, (mKeyCount - 1), 1, True
        End If
        decreaseKeyArray
        If mEmulateScriptingdictionary Then
            Dim objectKeyIndex : objectKeyIndex = getScriptingDictionaryObjectKeyIndex(key2)
            If objectKeyIndex > -1 Then removeDictionaryObjectKey objectKeyIndex
            If isScriptingDictionaryStoredKeyType(key2) Then mScriptingDictionaryStoredKeyTypes.Remove key2
        End If
    End If
End Function

' Removes all existing items and resets the DictCollection to array
Public Sub RemoveAll()
    initializeItemArray
    uninitializeKeyArray
    If mEmulateScriptingdictionary Then
        initializeDictionaryStoredKeyTypes
        initializeDictionaryObjectKeyArray
    Else
        uninitializeDictionaryStoredKeyTypes
        uninitializeDictionaryObjectKeyArray
    End If
End Sub

' Returns true if key exists, false of not
' variant key is compatible with Scripting.Dictionary
Public Function Exists(Key)
    Dim retVal 
    If mKeyCount > -1 Then
        If mLazySorting And (Not mKeysSorted) Then sortKeysAndRemoveDuplicates: mKeysSorted = True
        If mEmulateScriptingdictionary Then
            retVal = findKeyIndex(convertScriptingDictionaryToDictCollectionKey(Key)).Exact
        Else
            retVal = findKeyIndex(CStr(Key)).Exact
        End If
    End If
    Exists = retVal
End Function
' ======================== END: ITEM, ITEMAT, ADD, INSERT, ADD2, EXISTS, REMOVE, REMOVEALL ==========================================================



' ======================== START: COUNT, KEYCOUNT, INDEXOFKEY, KEYOFITEMAT, KEY =====================================================================
' returns the number of stored items (including Empty ones)
Public Property Get Count()
    ' sort keys and remove duplicate keys + their items, then count afterwards
    If mKeyCount > 1 And mLazySorting And (Not mKeysSorted) Then sortKeysAndRemoveDuplicates: mKeysSorted = True
    if mItemCount = -1 then Count = 0 else Count = mItemCount
End Property

' returns the number of stored keys
Public Property Get KeyCount()
    ' sort keys and remove duplicate keys + their items, then count afterwards
    If mKeyCount > 1 And mLazySorting And (Not mKeysSorted) Then sortKeysAndRemoveDuplicates: mKeysSorted = True
    if mKeyCount = -1 then KeyCount = 0 else KeyCount = mKeyCount 
End Property

' Returns index of given key. If ThrowErrors=False, returns -1 if key was not found
Public Function IndexOfKey(Key)
    Dim retVal : retVal = -1
    If mKeyCount < 0 Then
        ' is array or nonexisting or keys are empty
        If mThrowErrors Then Err.Raise 5, "DictCollection", "Invalid procedure call or argument"
        IndexOfKey = retVal: Exit Function 
    End If
    Dim result
    set result = new EvaluateKeyResult
    If mEmulateScriptingdictionary Then
        result.Key = convertScriptingDictionaryToDictCollectionKey(Key)
    ElseIf mEmulateCollection Then
        result.Key = convertCollectionToDictCollectionKey(Key)
    Else
        set result = evaluateKey(Key)
        If Not result.WasNumber And Not result.WasString Then
            If mThrowErrors Then Err.Raise 13, "DictCollection", "Keys of type " & TypeName(Key) & " are not supported"
            IndexOfKey = retVal: Exit Function 
        End If
        If mThrowErrors And (IsEmpty(Key) Or Len(result.Key) = 0) Then
            Err.Raise 5, "DictCollection", "Empty or zero length keys are not supported"
            IndexOfKey = retVal: Exit Function 
        end if
    End If
    If mKeyCount > 1 And mLazySorting And (Not mKeysSorted) Then sortKeysAndRemoveDuplicates: mKeysSorted = True
    Dim fkr
    set fkr = new FindKeyResult
    set fkr = findKeyIndex(result.Key)
    If fkr.Exact Then
        Dim itemIndex 
        itemIndex = allKeyItemIndexes(fkr.Index)
        If itemIndex < 0 Or itemIndex > (mItemCount - 1) Then
            WscripT.echo "!!!!!! Bug found in DictCollection function IndexOfKey(" & Key & "): key points to index=" & itemIndex & " outside of item array !!!!!!!"
            If mThrowErrors Then Err.Raise 9, "DictCollection", "Subscript out of range": Exit Function
        End If
        retVal = itemIndex
    Else
        If mThrowErrors Then Err.Raise 5, "DictCollection", "Key does not exist"
        IndexOfKey = retVal: Exit Function
    End If
    IndexOfKey = retVal
End Function

' Returns key of given item index. If ThrowErrors=False, returns empty string if key was not found
Public Property Get KeyOfItemAt(Index)
    KeyOfItemAt = GetMissingValue(, 1)
    If mKeyCount < 1 Then Exit Property
    If mItemCount < 1 Then Exit Property
    Dim internalIndex: internalIndex = Index - mExternalIndexOffset
    If internalIndex < 0 Or internalIndex > (mItemCount - 1) Then
        If mThrowErrors Then Err.Raise 9, "DictCollection", "No item at index " & Index
        Exit Property
    End If
    Dim keyIndex
    keyIndex = allItemKeyIndexes(internalIndex)
    If keyIndex < 0 Or keyIndex > (mKeyCount - 1) Then
        If mThrowErrors Then Err.Raise 9, "DictCollection", "No key at index " & keyIndex
        Exit Property
    End If
    Dim Key: Key = allKeys(keyIndex)
    If mEmulateScriptingDictionary Then
        If isStoredKeyOfTypeObject(Key) Then Set KeyOfItemAt = convertDictCollectionToScriptingDictionaryKey(Key) Else KeyOfItemAt = convertDictCollectionToScriptingDictionaryKey(Key)
    ElseIf mEmulateCollection Then
        KeyOfItemAt = convertDictCollectionToCollectionKey(Key)
    Else
        KeyOfItemAt = Key
    End If
End Property

' Changes an existing key to another value, if the new key value exists and ThrowErrors=False, its item will be removed
Public Property Let Key(Previous, Value)
    Dim previous2, value2 
    If mEmulateScriptingDictionary Then
        If IsObject(Previous) And IsObject(Value) Then
            If Previous Is Value Then Exit Property ' two objects are equal, exit
        ElseIf IsNumeric(Previous) And IsNumeric(Value) And VarType(Previous) <> VarType(Value) Then
            If Previous = Value Then Exit Property ' two numbers are equal, exit
        ElseIf IsDate(Previous) And IsDate(Value) Then
            If Previous = Value Then Exit Property ' two numbers are equal, exit
        ElseIf (IsNumeric(Previous) And VarType(Value) = vbDate) Or (VarType(Previous) = vbDate And IsNumeric(Value)) Then
            ' since date is not numeric, we have to check this differently
            If CDbl(Previous) = CDbl(Value) Then Exit Property ' number and date are equal, exit
        End If        
    Else
        If (VarType(Previous) <> vbString) Then
            ' emulate VBA.Collection behavior
            If mThrowErrors And Not mEmulateScriptingDictionary Then Err.Raise 13, "DictCollection", "Type mismatch": Exit Property
        End If
        If (VarType(Value) <> vbString) Then
            If IsEmpty(Value) Then
                ' emulate VBA.Collection behavior
                If mThrowErrors And Not mEmulateScriptingDictionary Then Err.Raise 13, "DictCollection", "Type mismatch": Exit Property
            End If
        End If
        If CStr(Previous) = "" Then
            ' since we don't support empty keys, emulate Scripting.Dictionary behavior for nonexisting keys
            If mThrowErrors Then Err.Raise 32811, "DictCollection", "Method 'Key' of 'DictCollection' failed."
            Exit Property
        ElseIf StrComp(CStr(Previous), CStr(Value), mCompareMode) = 0 Then
            ' two keys are equal, exit
            Exit Property
        End If
        If mKeyCount < 1 Or mItemCount < 1 Then
            If mThrowErrors Then Err.Raise 5, "DictCollection", "Invalid procedure call or argument"
            Exit Property
        End If
    End If
    
    If mKeyCount > 1 And mLazySorting And (Not mKeysSorted) Then sortKeysAndRemoveDuplicates: mKeysSorted = True
    Dim foundOldKey
    Set foundOldKey = New FindKeyResult
    ' find index of previous key
    If mEmulateScriptingDictionary Then
        previous2 = convertScriptingDictionaryToDictCollectionKey(Previous)
    ElseIf mEmulateCollection Then
        previous2 = convertCollectionToDictCollectionKey(Previous)
    Else
        previous2 = CStr(Previous)
    End If
    Set foundOldKey = findKeyIndex(previous2)
    
    If Not foundOldKey.Exact Then
        ' previous key not found, emulate Scripting.Dictionary behavior for nonexisting keys
        If mThrowErrors Then Err.Raise 32811, "DictCollection", "Method 'Key' of 'DictCollection' failed."
        Exit Property
    End If
    
    ' if we got here, we know that (1) previous key exists and (2) key / item arrays are not empty and (3) the new key is different from the old one
    Dim oldItemIndex : oldItemIndex = allKeyItemIndexes(foundOldKey.Index)
    If mEmulateScriptingDictionary Then
        value2 = convertScriptingDictionaryToDictCollectionKey(Value)
    Else
        value2 = CStr(Value)
    End If
    If value2 = "" Then
        ' new key is empty -> remove found previous (old) key
        If foundOldKey.Index = (mItemCount - 1) Then
            ' last item, do nothing
        Else
            copyKeysDown foundOldKey.Index + 1, (mKeyCount - 1), 1, True ' copy keys one down
        End If
        decreaseKeyArray
        ' and set keyIndex for corresponding item to -1 (= no key)
        allItemKeyIndexes(oldItemIndex) = -1
    Else
        Dim foundNewKey: Set foundNewKey = New FindKeyResult
        ' replace found previous (old) key
        set foundNewKey = findKeyIndex(value2)
        If foundNewKey.Exact Then
            ' new key exists!
            ' Emulate Scripting.Dictionary behavior
            If mThrowErrors Then Err.Raise 457, "DictCollection", "This key is already associated with an element of this collection": Exit Property
            Dim newItemIndex : newItemIndex = allKeyItemIndexes(foundNewKey.Index)
            ' (1) remove old key and adjust foundNewKey.index if necessary (if key was moved down)
            copyKeysDown foundOldKey.Index + 1, (mKeyCount - 1), 1, True ' copy keys one down
            If foundNewKey.Index > foundOldKey.Index Then foundNewKey.Index = foundNewKey.Index - 1
            allKeyItemIndexes(foundNewKey.Index) = oldItemIndex
            decreaseKeyArray
            ' (2) replace key of item with old key with new key
            allItemKeyIndexes(oldItemIndex) = foundNewKey.Index
            ' (3) remove item with new key (where new key previously pointed to)
            copyItemsDown newItemIndex + 1, (mItemCount - 1), 1, True
            decreaseItemArray
        Else
            ' new key not found -> foundNewKey.index is the index where the new key should be inserted
            If (foundOldKey.Index = foundNewKey.Index) Or (foundOldKey.Index = foundNewKey.Index - 1) Then
                ' new key should be inserted at the same index as old one -> new key overwrites old one
                allKeys(foundOldKey.Index) = value2
            ElseIf foundOldKey.Index > foundNewKey.Index Then
                ' new key should be inserted before old key -> copy all keys from new key index to old key index -1 one up
                copyKeysUp foundNewKey.Index, foundOldKey.Index - 1, 1, True
                allKeys(foundNewKey.Index) = value2
                allKeyItemIndexes(foundNewKey.Index) = oldItemIndex
                allItemKeyIndexes(oldItemIndex) = foundNewKey.Index
            ElseIf foundOldKey.Index < foundNewKey.Index Then
                ' new key should be inserted after old key -> copy all keys between old and new key one down
                copyKeysDown foundOldKey.Index + 1, foundNewKey.Index - 1, 1, True
                ' new key's index will be foundNewKey.index - 1 since old key is now gone
                allKeys(foundNewKey.Index - 1) = value2
                allKeyItemIndexes(foundNewKey.Index - 1) = oldItemIndex
                allItemKeyIndexes(oldItemIndex) = foundNewKey.Index - 1
            End If
            If mEmulateScriptingDictionary Then
                ' check if old key was object and if yes, remove it
                Dim oldObjectKeyIndex, objectKeyIndex : oldObjectKeyIndex = getScriptingDictionaryObjectKeyIndex(previous2)
                If oldObjectKeyIndex > -1 Then removeDictionaryObjectKey oldObjectKeyIndex
                ' check if new key is object and if yes, add it
                Dim newObjectKeyIndex : newObjectKeyIndex = getScriptingDictionaryObjectKeyIndex(value2)
                If newObjectKeyIndex > -1 Then UtilAddArrayValue mScriptingDictionaryObjectKeys, Value
                ' check if old key is number type and if yes, remove it
                If isScriptingDictionaryStoredKeyType(previous2) Then mScriptingDictionaryStoredKeyTypes.Remove previous2
                ' check if new key is number type and if yes, remember it
                If isScriptingDictionaryStoredKeyType(value2) Then mScriptingDictionaryStoredKeyTypes.Item(value2) = TypeName(Value)
            End If
        End If
    End If
End Property
' ======================== END: COUNT, KEYCOUNT, INDEXOFKEY, KEYOFITEMAT, KEY =======================================================================


' ======================== START: KEYS, SORTEDKEYS, INTERNALKEYS, ITEMS, INTERNALITEMS ==============================================================
' Returns all keys  array in the order of the items (as they were added)
' Variant array is compatible with Scripting.Dictionary
' VBScript Change: properties that return arrays cannot have parameters -> removed IncludeEmptyValues
Public Property Get Keys()
    Dim retVal, keyIndex 
    If mKeyCount > 0 And mLazySorting And (Not mKeysSorted) Then sortKeysAndRemoveDuplicates: mKeysSorted = True
    If mItemCount < 1 Then
        retVal = Array()
    Else
        Dim i
        ReDim retVal(mItemCount - 1)  
        keyIndex = 0
        For i = 0 To (mItemCount - 1)
            If allItemKeyIndexes(i) >= 0 Then
                ' if item has key, copy key to output array
                If mEmulateScriptingdictionary Then
                    Dim Key : Key = allKeys(allItemKeyIndexes(i))
                    If isStoredKeyOfTypeObject(Key) Then Set retVal(keyIndex) = convertDictCollectionToScriptingDictionaryKey(Key) Else retVal(keyIndex) = convertDictCollectionToScriptingDictionaryKey(Key)
                ElseIf mEmulateCollection Then
                    retVal(keyIndex) = convertDictCollectionToCollectionKey(Key)
                Else
                    retVal(keyIndex) = allKeys(allItemKeyIndexes(i))
                End If
                keyIndex = keyIndex + 1
            Else
                ' if not, fill in missing value as Empty
                retVal(keyIndex) = Empty
                keyIndex = keyIndex + 1
            End If
        Next
    End If
    Keys = retVal
End Property

' Returns all keys  array in the inernal sorted order
' Variant array instead of String arrray because then it can also be used with Scripting.Dictionary emulation
Public Property Get SortedKeys() 
    Dim retVal 
    If mKeyCount > 0 And mLazySorting And (Not mKeysSorted) Then sortKeysAndRemoveDuplicates: mKeysSorted = True
    If mItemCount < 1 Or mKeyCount < 1 Then
        retVal = Array()
    Else
        Dim i 
        ReDim retVal(mKeyCount - 1)
        For i = 0 To (mKeyCount - 1)
            retVal(i) = allKeys(i)
            ' if item has key, copy key to output array
            If mEmulateScriptingdictionary Then
                Dim Key : Key = retVal(i)
                If isStoredKeyOfTypeObject(Key) Then Set retVal(i) = convertDictCollectionToScriptingDictionaryKey(Key) Else retVal(i) = convertDictCollectionToScriptingDictionaryKey(Key)
            End If
        Next
    End If
    SortedKeys = retVal
End Property

' returns a reference to the internal keys array
Public Property Get InternalKeys(): InternalKeys = allKeys: End Property

' Returns all items  array in the order they were added
' Variant array is compatible with Scripting.Dictionary
Public Property Get Items()
    Dim retVal 
    If mItemCount < 1 Then
        retVal = Array()
    Else
        Dim i
        ReDim retVal((mItemCount - 1))
        For i = 0 to (mItemCount - 1)
            If IsObject(allItems(i)) Then Set retVal(i) = allItems(i) Else retVal(i) = allItems(i)
        Next
    End If
    Items = retVal
End Property

' returns a reference to the internal items array
Public Property Get InternalItems(): InternalItems = allItems: End Property
' ======================== END: KEYS, SORTEDKEYS, INTERNALKEYS, ITEMS, INTERNALITEMS ================================================================


' ======================== START: INTERNAL HELPER FUNCTIONS - LOOKUP, SORTING, COPYING ==============================================================
' gets the item index from a Variant that can be both, index or key and remembers the source data type
Private Function evaluateIndexOrKey(IndexOrKey)
    Dim retVal: set retVal = new EvaluateIndexOrKeyResult
    Select Case TypeName(IndexOrKey)
        Case "String"
            retVal.WasKey = True
            If IndexOrKey = vbNullString Then
                'do not search for key
                retVal.Index = -1
            Else
                'a string was passed, search for key
                Dim Key, resultKey
                set resultKey = new FindKeyResult
                Key = IndexOrKey
                set resultKey = findKeyIndex(Key)
                If resultKey.Exact Then
                    'key was found
                    retVal.Index = allKeyItemIndexes(resultKey.Index)
                Else
                    'key was not found
                    retVal.Index = -1
                End If
            End If
        Case "Integer", "Long", "Byte"
            'an integer was passed, emulate Collection behavior by accepting all these types
            retVal.WasIndex = True
            retVal.Index = IndexOrKey - mExternalIndexOffset
        Case "Single", "Double", "Currency", "Decimal", "Date"
            'a fractional number was passed, convert to long
            retVal.WasIndex = True
            retVal.Index = CLng(IndexOrKey) - mExternalIndexOffset
        Case Else
            'argument is neither number nor string
            retVal.Index = -1
    End Select
    set evaluateIndexOrKey = retVal
End Function

' converts a Variant key to String and remembers the source data type
Private Function evaluateKey(Key)
    Dim retVal
    set retVal = new EvaluateKeyResult
    Select Case TypeName(Key)
        Case "Integer", "Long", "Single", "Double", "Currency", "Decimal", "Date":  retVal.WasNumber = True: retVal.Key = CStr(Key)
        Case "String":  retVal.WasString = True: retVal.Key = Key
    End Select
    set evaluateKey = retVal
End Function

' Returns index of key. Snaps to first stored key that matches the searched key (e.g. case insensitive search).
' If key is not found, returns
' 1) [-1] if searched key is less than first stored key
' 2) [last stored key index + 1] if searched key is greater than last stored key
Private Function findKeyIndex(Key)
    Dim upper, lower, row, counter, retVal, rowBefore
    set retVal = new FindKeyResult
    retVal.Exact = False
    If mItemCount < 1 Or mKeyCount < 1 Then
        retVal.Index = 0
    Else
        lower = 0:  upper = mKeyCount - 1: counter = 0:
        'compare with last key
        Select Case StrComp(Key, allKeys(upper), mCompareMode) 'compare with last element
            Case 0: retVal.Index = upper: retVal.Exact = True 'equals last key
            Case 1: retVal.Index = upper + 1 'greater than last key
            Case -1
                'compare with first key
                Select Case StrComp(Key, allKeys(lower), mCompareMode)
                    Case 0: retVal.Index = lower: retVal.Exact = True 'equals first key
                    Case -1: retVal.Index = lower 'less than first key
                    Case 1:
                        ' search using quick search
                        row = (upper - lower) \ 2
                        Do
                            counter = counter + 1
                            Select Case StrComp(Key, allKeys(row), mCompareMode)
                                Case 0:
                                    ' ensure that it snaps to first key that matches by going back until key before does not match
                                    rowBefore = row - 1
                                    Do While rowBefore > -1 And StrComp(Key, allKeys(rowBefore), mCompareMode) = 0
                                        row = rowBefore: rowBefore = row - 1
                                    Loop
                                    retVal.Index = row: retVal.Exact = True
                                Case 1: retVal.Index = row + 1: lower = row: row = lower + ((upper - lower) \ 2)
                                Case -1: retVal.Index = row: upper = row: row = lower + ((upper - lower) \ 2)
                            End Select
                        Loop While ((upper - lower) > 1) And (retVal.Exact = False) And counter < 7000
                        If counter = 7000 Then WscripT.echo "!!!!!! Bug found in DictCollection: Endless loop at findKeyIndex(""" & Key & """)!!!!!!!"
                End Select
        End Select
    End If
    set findKeyIndex = retVal
End Function

' performs key sorting and removes duplicates can exist when adding items and keys while LazySorting=True
Private Sub sortKeysAndRemoveDuplicates()
    If mKeyCount < 0 Then Exit Sub
    sortKeys 0, (mKeyCount - 1)
    removeKeyDuplicates
End Sub

' removes all duplicate key entries by running over all sorted keys and dropping all but the last key/item pairs while preserving the first item index for that key
Private Sub removeKeyDuplicates()
    Dim i, j, k, lastKey, itemIndexes, Offset, match, last, swappedValue 
    ' remove duplicate keys and corresponding items
    lastKey = allKeys(0): itemIndexes = Array(): i = 1
    While i <= UBound(allKeys)
        If StrComp(lastKey, allKeys(i), mCompareMode) = 0 Then
            match = True
            ' add item index to array
            If UBound(itemIndexes) = -1 Then
                ' add previous and current item indexes
                ReDim Preserve itemIndexes(1)
                itemIndexes(0) = allKeyItemIndexes(i - 1)
                itemIndexes(1) = allKeyItemIndexes(i)
            Else
                ' add current item index
                ReDim Preserve itemIndexes(UBound(itemIndexes) + 1)
                itemIndexes(UBound(itemIndexes)) = allKeyItemIndexes(i)
            End If
        Else
            match = False
        End If
        last = (i = UBound(allKeys))
        ' if adjacent keys do not match or if last key
        If last Or Not match Then
            Offset = UBound(itemIndexes)
            If Offset > 0 Then
                ' sort indexes
                If last Then i = i + 1
                UtilSortArray itemIndexes, LBound(itemIndexes), UBound(itemIndexes)
                
                ' swap first and last item to preserve item index
                swappedValue = allItems(itemIndexes(UBound(itemIndexes)))
                allItems(itemIndexes(UBound(itemIndexes))) = allItems(itemIndexes(0))
                allItems(itemIndexes(0)) = swappedValue
                
                ' set the last keys item index to the item with the (now) lowest index (last item added removes all others)
                allKeyItemIndexes(i - 1) = itemIndexes(0)
                ' remove all similar keys except last one
                copyKeysDown i - 1, UBound(allKeys), Offset, True
                ' set the overwriting items key to the remaining unique key
                allItemKeyIndexes(itemIndexes(0)) = (i - 1) - Offset
                
                ' set arrays to new size
                ReDim Preserve allKeys(UBound(allKeys) - Offset)
                ReDim Preserve allKeyItemIndexes(UBound(allKeyItemIndexes) - Offset)
                
                ' remove corresponding items, starting from last to second
                For j = UBound(itemIndexes) To 1 Step -1
                    ' move last index one down with each iteration to avoid adjusting keyItemIndexes of last items in array multiple times
                    copyItemsDown itemIndexes(j) + 1, UBound(allItems) - (UBound(itemIndexes) - j), 1, True
                Next
                ReDim Preserve allItems(UBound(allItems) - UBound(itemIndexes))
                ReDim Preserve allItemKeyIndexes(UBound(allItemKeyIndexes) - UBound(itemIndexes))

                ' reset variables
                itemIndexes = Array()
                i = i - Offset
            End If
            If Not last Then lastKey = allKeys(i)
        End If
        i = i + 1
    Wend
    mItemCount = UBound(allItems) + 1
    mKeyCount = UBound(allKeys) + 1
End Sub

' https://wellsr.com/vba/2018/excel/vba-quicksort-macro-to-sort-arrays-fast/
Private Sub sortKeys(ArrayLowerBound, ArrayUpperBound)
    ' StrComp(string1, string2) -> [string1 < string2] = -1, [string1 > string2] = 1, [string1 = string2] = 0
    Dim splitKey, swappedKey, swappedKeyItemIndex, swappedItemKeyIndex, lower, upper 
    lower = ArrayLowerBound
    upper = ArrayUpperBound
    splitKey = allKeys((ArrayLowerBound + ArrayUpperBound) \ 2)
    While (lower <= upper) 'divide
        While (StrComp(allKeys(lower), splitKey, mCompareMode) = -1 And lower < ArrayUpperBound)
           lower = lower + 1
        Wend
        While (StrComp(splitKey, allKeys(upper), mCompareMode) = -1 And upper > ArrayLowerBound)
           upper = upper - 1
        Wend
        If (lower <= upper) Then
            ' swap key
            swappedKey = allKeys(lower)
            allKeys(lower) = allKeys(upper)
            allKeys(upper) = swappedKey
            ' swap corresponding key-to-item pointer
            swappedKeyItemIndex = allKeyItemIndexes(lower)
            allKeyItemIndexes(lower) = allKeyItemIndexes(upper)
            allKeyItemIndexes(upper) = swappedKeyItemIndex
            ' swap item-to-key pointer
            swappedItemKeyIndex = allItemKeyIndexes(allKeyItemIndexes(lower))
            allItemKeyIndexes(allKeyItemIndexes(lower)) = allItemKeyIndexes(allKeyItemIndexes(upper))
            allItemKeyIndexes(allKeyItemIndexes(upper)) = swappedItemKeyIndex
            lower = lower + 1
            upper = upper - 1
        End If
    Wend
    If (ArrayLowerBound < upper) Then sortKeys ArrayLowerBound, upper 'conquer
    If (lower < ArrayUpperBound) Then sortKeys lower, ArrayUpperBound 'conquer
End Sub

' copies items to higher index by offset and increases their key-to-item pointers if needed (to reflect the new item indexes)
' will copy only the items that can be copied without changing the array size
Private Sub copyItemsUp(FirstIndex, LastIndex, Offset, IncreaseTheirKeyItemIndexByOffset)
    Dim sourceIndex, keyIndex, lower, upper 
    ' ensure that copying is possible within array bounds
    If Offset >= mItemCount Then Exit Sub
    If FirstIndex < 0 Then lower = 0 Else lower = FirstIndex
    If LastIndex > (mItemCount - 1 - Offset) Then upper = mItemCount - 1 - Offset Else upper = LastIndex
    For sourceIndex = upper To lower Step -1
        If IsObject(allItems(sourceIndex)) Then Set allItems(sourceIndex + Offset) = allItems(sourceIndex) Else allItems(sourceIndex + Offset) = allItems(sourceIndex)
        keyIndex = allItemKeyIndexes(sourceIndex)
        allItemKeyIndexes(sourceIndex + Offset) = keyIndex
        If IncreaseTheirKeyItemIndexByOffset Then
            If keyIndex > -1 And keyIndex < mKeyCount Then allKeyItemIndexes(keyIndex) = allKeyItemIndexes(keyIndex) + Offset
        End If
    Next
End Sub

' copies items to lower index by offset and decreases their key-to-item pointers if needed (to reflect the new item indexes)
' will copy only the items that can be copied without changing the array size
Private Sub copyItemsDown(FirstIndex, LastIndex, Offset, DecreaseTheirKeyItemIndexByOffset)
    Dim sourceIndex, keyIndex, lower, upper 
    ' ensure that copying is possible within array bounds
    If Offset >= mItemCount Then Exit Sub
    If (FirstIndex - Offset) < 0 Then lower = Offset Else lower = FirstIndex
    If LastIndex >= mItemCount Then upper = mItemCount - 1 Else upper = LastIndex
    For sourceIndex = lower To upper Step 1
        If IsObject(allItems(sourceIndex)) Then Set allItems(sourceIndex - Offset) = allItems(sourceIndex) Else allItems(sourceIndex - Offset) = allItems(sourceIndex)
        keyIndex = allItemKeyIndexes(sourceIndex)
        allItemKeyIndexes(sourceIndex - Offset) = keyIndex
        If DecreaseTheirKeyItemIndexByOffset Then
            If keyIndex > -1 And keyIndex < mKeyCount Then allKeyItemIndexes(keyIndex) = allKeyItemIndexes(keyIndex) - Offset
        End If
    Next
End Sub

' copies keys to lower index by offset and increases their item-to-key pointers if needed (to reflect the new key indexes)
' will copy only the keys that can be copied without changing the array size
Private Sub copyKeysUp(FirstIndex, LastIndex, Offset, IncreaseTheirItemKeyIndexByOffset)
    Dim sourceIndex, itemIndex, lower, upper 
    ' ensure that copying is possible within array bounds
    If Offset >= mKeyCount Then Exit Sub
    If FirstIndex < 0 Then lower = 0 Else lower = FirstIndex
    If LastIndex > (mKeyCount - 1 - Offset) Then upper = mKeyCount - 1 - Offset Else upper = LastIndex
    For sourceIndex = upper To lower Step -1
        allKeys(sourceIndex + Offset) = allKeys(sourceIndex)
        itemIndex = allKeyItemIndexes(sourceIndex)
        allKeyItemIndexes(sourceIndex + Offset) = itemIndex
        If IncreaseTheirItemKeyIndexByOffset Then
            If itemIndex > -1 And itemIndex < mItemCount Then allItemKeyIndexes(itemIndex) = allItemKeyIndexes(itemIndex) + Offset
        End If
    Next
End Sub

' copies keys to higher index by offset and decreases their item-to-key pointers if needed (to reflect the new key indexes)
' will copy only the keys that can be copied without changing the array size
Private Sub copyKeysDown(FirstIndex, LastIndex, Offset, DecreaseTheirItemKeyIndexByOffset)
    Dim sourceIndex, itemIndex, lower, upper 
    ' ensure that copying is possible within array bounds
    If Offset >= mKeyCount Then Exit Sub
    if (FirstIndex - Offset) < 0 then lower = Offset else lower = FirstIndex
    if LastIndex >= mKeyCount then upper = mKeyCount - 1 else upper = LastIndex
    For sourceIndex = lower To upper Step 1
        allKeys(sourceIndex - Offset) = allKeys(sourceIndex)
        itemIndex = allKeyItemIndexes(sourceIndex)
        allKeyItemIndexes(sourceIndex - Offset) = itemIndex
        If DecreaseTheirItemKeyIndexByOffset Then
            If itemIndex > -1 And itemIndex < mItemCount Then allItemKeyIndexes(itemIndex) = allItemKeyIndexes(itemIndex) - Offset
        End If
    Next
End Sub
' ======================== END: INTERNAL HELPER FUNCTIONS - LOOKUP, SORTING, COPYING ================================================================




' ======================== START: EMULATION HELPER FUNCTIONS =========================================================================================
' used by Scripting.Dictionary compatibility: converts Dictionary Variant keys to String keys that can be stored by DictCollection
Function convertScriptingDictionaryToDictCollectionKey(IndexOrKey) 
    Dim Key, tname 
    tname = TypeName(IndexOrKey)
    Select Case tname
        Case "String": If IndexOrKey = "" Then Key = "[EMP:]" Else Key = IndexOrKey
        Case "Integer", "Long", "Single", "Double", "Currency", "Decimal", "Date":
            Key = "[NUM:" & CStr(IndexOrKey) & "]"
        Case "Byte":
            Dim byteArray(): byteArray = IndexOrKey
            Key = "[BTE:" & StrConv(byteArray, vbUnicode) & "]"
        Case "Boolean": Key = "[BOL:" & CStr(IndexOrKey) & "]"
        Case "Error":
            Dim ErrNumber : ErrNumber = CStr(IndexOrKey)
            ' remove the "Error " part
            Key = "[ERR:" & Mid(ErrNumber, 7, Len(ErrNumber) - 6) & "]"
        Case "Null": Key = "[NUL:]"
        Case "Empty": Key = "[EMP:]"
        Case Else:
            If IsObject(IndexOrKey) Then
                ' find object in array. if not found, set key to index after last index
                Dim Index : Index = UtilFindArrayIndex(mScriptingDictionaryObjectKeys, IndexOrKey)
                If Index > -1 Then Key = "[OBJ:" & Index & "]" Else Key = "[OBJ:" & UBound(mScriptingDictionaryObjectKeys) + 1 & "]"
            End If
    End Select
    convertScriptingDictionaryToDictCollectionKey = Key
End Function

' used by Scripting.Dictionary compatibility: converts a stored String key back to Scripting.Dictionary Variant (String, Number, Object, Error or Empty)
Private Function convertDictCollectionToScriptingDictionaryKey(Key)
    Dim keyType, keyVal, originalDataType, retVal 
    retVal = Key
    If Len(Key) >= 6 Then
        keyType = Mid(Key, 2, 3): keyVal = Mid(Key, 6, Len(Key) - 6)
        Select Case keyType
            Case "NUM":
                originalDataType = mScriptingDictionaryStoredKeyTypes.Item(Key)
                Select Case originalDataType
                    Case "Integer": retVal = CInt(keyVal)
                    Case "Long": retVal = CLng(keyVal)
                    Case "Single": retVal = CSng(keyVal)
                    Case "Double": retVal = CDbl(keyVal)
                    Case "Currency": retVal = CCur(keyVal)
                    Case "Decimal": retVal = CDec(keyVal)
                    Case "Date": retVal = CDate(keyVal)
                End Select
            Case "BTE": retVal = StrConv(keyVal, vbFromUnicode)
            Case "BOL": retVal = CBool(keyVal)
            Case "OBJ":
                Dim objectKeyIndex : objectKeyIndex = getScriptingDictionaryObjectKeyIndex(Key)
                If objectKeyIndex > -1 Then
                    Set retVal = mScriptingDictionaryObjectKeys(objectKeyIndex)
                Else
                    WscripT.echo "!!!!!! Bug found in DictCollection property Keys: objectKeyIndex=-1 !!!!!!!"
                    Set retVal = GetMissingValue(, 1)
                End If
            Case "ERR": retVal = CVErr(CLng(keyVal))
            Case "EMP"
                originalDataType = mScriptingDictionaryStoredKeyTypes.Item(Key)
                Select Case originalDataType
                    Case "String": retVal = ""
                    Case "Empty": retVal = Empty
                End Select
        End Select
    End If
    If IsObject(retVal) Then Set convertDictCollectionToScriptingDictionaryKey = retVal Else convertDictCollectionToScriptingDictionaryKey = retVal
End Function

' used by Scripting.Dictionary compatibility: returns true if stored key is an object (needed to distinguish between set/let syntax)
Private Function isStoredKeyOfTypeObject(Key)
    isStoredKeyOfTypeObject = False
    If Len(Key) >= 6 Then
        If Mid(Key, 2, 3) = "OBJ" And Right(Key, 1) = "]" Then isStoredKeyOfTypeObject = True
    End If
End Function

' used by Scripting.Dictionary compatibility: checks if stored key has a data type that is stored mScriptingDictionaryStoredKeyTypes:
' Integer, Long, Single, Double, Currency, Decimal, etc.
Private Function isScriptingDictionaryStoredKeyType(Key) 
    isScriptingDictionaryStoredKeyType = False
    If Len(Key) > 4 Then
        Select Case Left(Key, 4)
            Case "[NUM", "[EMP":
                isScriptingDictionaryStoredKeyType = True
        End Select
    End If
End Function

' used by Scripting.Dictionary compatibility: stores the key object or the original number datatype if necessary
Private Sub storeScriptingDictionaryKeyInformation(ExternalKey, InternalKey)
     Dim objectKeyIndex
     ' check if key was object and add it if it does not exist or re-set it, if it does exist
     objectKeyIndex = getScriptingDictionaryObjectKeyIndex(InternalKey)
     If UBound(mScriptingDictionaryObjectKeys) > -1 Or objectKeyIndex > -1 Then
         If objectKeyIndex > UBound(mScriptingDictionaryObjectKeys) Then UtilAddArrayValue mScriptingDictionaryObjectKeys, ExternalKey
    End If
     ' check if key needs to be stored and if yes, store original key type name for backward conversion
     If isScriptingDictionaryStoredKeyType(InternalKey) Then mScriptingDictionaryStoredKeyTypes.Item(InternalKey) = TypeName(ExternalKey)
End Sub

' used by Scripting.Dictionary compatibility: checks if key is object and if yes, extracts index from key. if not, returns -1
Private Function getScriptingDictionaryObjectKeyIndex(Key) 
    Dim retVal : retVal = -1
    If Len(Key) > 5 Then
        If Left(Key, 5) = "[OBJ:" Then retVal = CLng(Mid(Key, 6, Len(Key) - 6))
    End If
    getScriptingDictionaryObjectKeyIndex = retVal
End Function

' used by Scripting.Dictionary compatibility: removes an object from the object keys array and decreases all object key references greater/equal index in allKeys by one
Private Sub removeDictionaryObjectKey(Index)
    Dim i, objectKeyIndex 
    ' remove old object key
    UtilRemoveArrayValueByIndex mScriptingDictionaryObjectKeys, Index
    ' run over all keys and decrease the existing object key indexes by 1 if they are greater than the old object key index
    For i = 0 to (mKeyCount - 1)
        objectKeyIndex = getScriptingDictionaryObjectKeyIndex(allKeys(i))
        If objectKeyIndex >= Index Then allKeys(i) = "[OBJ:" & objectKeyIndex - 1 & "]"
    Next
End Sub

' Used by Collection compatibility: maps Collection Keys (Zero Length Strings and Missing) to DictCollection compatibe IndexOrKey
Function convertCollectionToDictCollectionKey(IndexOrKey) 
    Dim key2, tname 
    If IsMissing(IndexOrKey) Then
        key2 = ""
    Else
        tname = TypeName(IndexOrKey)
        Select Case tname
            Case "String"
                If IndexOrKey = "" Then key2 = "[EMP:]" Else key2 = IndexOrKey
            Case Else
                key2 = IndexOrKey
        End Select
    End If
    convertCollectionToDictCollectionKey = key2
End Function

' used by Collection compatibility: converts a stored String key back to Collection Variant (String or Empty)
Private Function convertDictCollectionToCollectionKey(Key) 
    Dim keyType, keyVal, retVal 
    retVal = Key
    If Len(Key) >= 6 Then
        keyType = Mid(Key, 2, 3): keyVal = Mid(Key, 6, Len(Key) - 6)
        Select Case keyType
            Case "EMP": retVal = ""
        End Select
    End If
    convertDictCollectionToScriptingDictionaryKey = retVal
End Function
' ======================== END: EMULATION HELPER FUNCTION ===========================================================================================



' ======================== START: EXTENDED FUNCTIONALITY ============================================================================================
' adds a new DictCollection by given index or key and returns it
' will copy parent DictCollection Settings
Public Function AddDictCollection(IndexOrKey)
    Dim newDictCollection
    Set newDictCollection = new DictCollection
    internalSetItem IndexOrKey, newDictCollection
    Set AddCollection = newDictCollection
End Function

' retrieves a DictCollection from a given index or key; adds a DictCollection if the item is nonexisting or not a DictCollection
Public Function GetOrSetDictCollection(IndexOrKey)
    Dim newOrExistingDictCollection, result, action  ' 0 = found, 1 = found/overwrite with new, 2 = not found/add
    Set result = new EvaluateIndexOrKeyResult
    If mItemCount = -1 Then
        initializeItemArray
        action = 2 ' not found/add
    Else
        If mItemCount < 1 Then
            action = 2 ' not found/add
        Else
            If mKeyCount > 1 And mLazySorting And (Not mKeysSorted) Then sortKeysAndRemoveDuplicates: mKeysSorted = True
            set result = evaluateIndexOrKey(IndexOrKey)
            If result.Index < 0 Or result.Index > (mItemCount - 1) Then
                action = 2 ' not found/add
            Else
                ' found
                If IsObject(allItems(result.Index)) Then
                    If InStr(TypeName(allItems(result.Index)), "DictCollection") > 0 Then action = 0 Else action = 1 ' found/overwrite with new
                Else
                    action = 1 ' found/overwrite with new
                End If
            End If
        End If
    end if
    Select Case action
        Case 0: Set newOrExistingDictCollection = allItems(result.Index)
        Case 1, 2:
            Set newOrExistingDictCollection = new DictCollection
            CopyDefaultValueSettings Me, newOrExistingDictCollection
            internalSetItem IndexOrKey, newOrExistingDictCollection
    End Select
    Set GetOrSetDictCollection = newOrExistingDictCollection
End Function
' ======================== END: EXTENDED FUNCTIONALITY ============================================================================================



' ======================== START: PUBLIC HELPER FUNCTIONS =========================================================================================
' Sorts a one-dimensional VBA array from smallest to largest using a very fast quicksort algorithm
' source: https://wellsr.com/vba/2018/excel/vba-quicksort-macro-to-sort-arrays-fast/
Public Sub UtilSortArray(arr, ArrayLowerBound, ArrayUpperBound)
    Dim splitValue, swappedValue, lower, upper 
    lower = ArrayLowerBound
    upper = ArrayUpperBound
    splitValue = arr((ArrayLowerBound + ArrayUpperBound) \ 2)
    While (lower <= upper) 'divide
       While (arr(lower) < splitValue And lower < ArrayUpperBound)
          lower = lower + 1
       Wend
       While (splitValue < arr(upper) And upper > ArrayLowerBound)
          upper = upper - 1
       Wend
       If (lower <= upper) Then
          swappedValue = arr(lower)
          arr(lower) = arr(upper)
          arr(upper) = swappedValue
          lower = lower + 1
          upper = upper - 1
       End If
    Wend
  If (ArrayLowerBound < upper) Then UtilSortArray arr, ArrayLowerBound, upper 'conquer
  If (lower < ArrayUpperBound) Then UtilSortArray arr, lower, ArrayUpperBound 'conquer
End Sub

' randomizes the positions of the values in an array without changing them
Public Sub UtilShuffleArray(arr)
    Dim oldIndex, newIndex, swappedValue 
    Randomize
    For oldIndex = LBound(arr) To UBound(arr)
        newIndex = CLng(((UBound(arr) - oldIndex) * Rnd) + oldIndex) ' move value somewhere upwards in array
        If oldIndex <> newIndex Then swappedValue = arr(oldIndex): arr(oldIndex) = arr(newIndex): arr(newIndex) = swappedValue
    Next
End Sub

' adds a value to an array
Public Sub UtilAddArrayValue(arr, val)
    ' initialize array if arr is not an array
    If InStr(TypeName(arr), "()") < 1 Then arr = Array() Else ReDim Preserve arr(UBound(arr) + 1)
    ' set object reference or value
    If IsObject(val) Then Set arr(UBound(arr)) = val Else arr(UBound(arr)) = val
End Sub

' removes all occurrences of a value or object reference from an array
Public Sub UtilRemoveArrayValue(arr, val)
    Dim i, newArr, newArrIndex 
    If InStr(TypeName(arr), "()") < 1 Then arr = Array(): Exit Sub
    ReDim newArr(UBound(arr)): newArrIndex = 0
    For i = LBound(arr) To UBound(arr)
        ' copy object reference or value if not equal
        If IsObject(arr(i)) And IsObject(val) Then
            If Not (arr(i) Is val) Then
                Set newArr(newArrIndex) = arr(i): newArrIndex = newArrIndex + 1
            Else
                ReDim Preserve newArr(UBound(newArr) - 1)
            End If
        Else
            If arr(i) <> val Then
                newArr(newArrIndex) = arr(i): newArrIndex = newArrIndex + 1
            Else
                ReDim Preserve newArr(UBound(newArr) - 1)
            End If
        End If
    Next
    arr = newArr
End Sub

' removes an array value from the array at a given index
Public Sub UtilRemoveArrayValueByIndex(arr, Index)
    Dim i 
    If InStr(TypeName(arr), "()") < 1 Or (Index = 0 And (UBound(arr) = LBound(arr))) Then arr = Array(): Exit Sub
    If Index < LBound(arr) Or Index > UBound(arr) Then Exit Sub
    If Index < UBound(arr) Then
        ' copy all values above index one down
        For i = Index + 1 To UBound(arr)
            If IsObject(arr(i)) Then Set arr(i - 1) = arr(i) Else arr(i - 1) = arr(i)
        Next
    End If
    ' decrease array length by 1
    ReDim Preserve arr(UBound(arr) - 1)
End Sub

' finds first occurrence of a value or object refernce within an array or returns -1 if not found
Public Function UtilFindArrayIndex(arr, val) 
    Dim i 
    If InStr(TypeName(arr), "()") < 1 Then arr = Array(): UtilFindArrayIndex = -1: Exit Function
    For i = LBound(arr) To UBound(arr)
        ' copy object reference or value if not equal
        If IsObject(arr(i)) And IsObject(val) Then
            If (arr(i) Is val) Then UtilFindArrayIndex = i: Exit Function
        Else
            If arr(i) = val Then UtilFindArrayIndex = i: Exit Function
        End If
    Next
    UtilFindArrayIndex = -1
End Function

Public Function GetMissingValue(DoNotPassAnythingHere, LastArgument) : GetMissingValue = DoNotPassAnythingHere: End Function
' ======================== END: PUBLIC HELPER FUNCTIONS ===========================================================================================


' ======================== START: DEMO AND SELFTEST ===============================================================================================
' performs demo of basic functions
Public Sub DemoBasicFunctionality()
    Dim dc: Set dc = New DictCollection
    wscript.echo "----------- Start DictCollection Demo of Basic Functionality -------------------"
    wscript.echo "  dc.Add ""key1"", 123.45": dc.Add "key1", 123.45
    wscript.echo "     dc.Item(0) should be 123.45 -> " & dc.Item(0)
    wscript.echo "  dc.Add, ""A""": dc.Add, "A"
    wscript.echo "     dc.Item(1) should be A -> " & dc.Item(1)
    wscript.echo "     dc.Count should be 2 -> " & dc.Count
    wscript.echo "     dc.KeyCount should be 1 -> " & dc.KeyCount
    wscript.echo "     dc.Item(""key1"") should be 123.45 -> " & dc.Item("key1")
    wscript.echo "     dc.Item(1) should be A -> " & dc.Item(1)
    wscript.echo "  dc.Item(0) = 100": dc.Item(0) =100
    wscript.echo "     dc.Item(0) should be 100 -> " & dc.Item(0)
    wscript.echo "  dc.Item(""key1"") = 200": dc.Item("key1") = 200
    wscript.echo "     dc.Item(0) should be 200 -> " & dc.Item(0)
    wscript.echo "  dc.Item(""key2"") = ""B""": dc.Item("key2") = "B"
    wscript.echo "     dc.Item(2) should be B -> " & dc.Item(2)
    wscript.echo "  dc.Add ""key3"", New DictCollection": dc.Add "key3", New DictCollection
    wscript.echo "     TypeName(dc.Item(3)) should be DictCollection -> " & TypeName(dc.Item(3))
    wscript.echo "  dc.Item(""key3"").Add ""key4"", 456.78""": dc.Item("key3").Add "key4", 456.78
    wscript.echo "     dc.Item(3).Item(0) should be 456.78 -> " & dc.Item(3).Item(0)
    wscript.echo "  dc.Item(""key3"").Add, ""C""": dc.Item("key3").Add, "C"
    wscript.echo "     dc.Item(3).Item(1) should be C -> " & dc.Item(3).Item(1)
    wscript.echo "     dc.Item(""key3"").Item(""key4"") should be 456.78 -> " & dc.Item("key3").Item("key4")
    wscript.echo "     dc.Item(""key3"").Item(1) should be C -> " & dc.Item("key3").Item(1)
    wscript.echo "     (dc.Item(""key3"").Item(1)=dc.NonExistingValue) should be False -> " & (dc.Item("key3").Item(1) = dc.NonExistingValue)
    wscript.echo "     (dc.Item(""key3"").Item(2)=dc.NonExistingValue) should be True -> " & (dc.Item("key3").Item(2) = dc.NonExistingValue)
    wscript.echo "  dc.Remove(""key1"")": dc.Remove("key1")
    wscript.echo "     dc.Item(0) should be A -> " & dc.Item(0)
    wscript.echo "  dc.Remove(0)": dc.Remove(0)
    wscript.echo "     dc.Item(0) should be B -> " & dc.Item(0)
    wscript.echo "     dc.KeyOfItemAt(0) should be key2 -> " & dc.KeyOfItemAt(0)
    wscript.echo "  dc.Key(""key2"") = ""key4""": dc.Key("key2") = "key4"
    wscript.echo "     dc.KeyOfItemAt(0) should be key4 -> " & dc.KeyOfItemAt(0)
    wscript.echo "     dc.Exists(""key3"") should be True -> " & dc.Exists("key3")
    wscript.echo "     dc.Count should be 2 -> " & dc.Count
    wscript.echo "  dc.Key(""key4"") = ""key3""": dc.Key("key4") = "key3"
    wscript.echo "     dc.KeyOfItemAt(0) should be key3 -> " & dc.KeyOfItemAt(0)
    wscript.echo "     dc.Exists(""key4"") should be False -> " & dc.Exists("key4")
    wscript.echo "     dc.Count should be 1 -> " & dc.Count
    wscript.echo "     Join(dc.Keys,"","") should be key3 -> " & Join(dc.Keys,",")
    wscript.echo "     Join(dc.Items,"","") should be B -> " & Join(dc.Items,",")
    wscript.echo "  dc.Insert ""D"", 0, ""key5""": dc.Insert "D", 0, "key5"
    wscript.echo "     dc.Item(0) should be D -> " & dc.Item(0)
    wscript.echo "     dc.Item(1) should be B -> " & dc.Item(1)
    wscript.echo "  dc.Insert ""E"", 3, """"": dc.Insert "E", 3, ""
    wscript.echo "     Join(dc.Items,"","") should be D,B,,E -> " & Join(dc.Items,",")
    wscript.echo "     IsEmpty(dc.Item(2)) should be True -> " & IsEmpty(dc.Item(2))
    wscript.echo "     Join(dc.Keys,"","") should be key5,key3,, -> " & Join(dc.Keys,",")
    wscript.echo "     Join(dc.SortedKeys,"","") should be key3,key5 -> " & Join(dc.SortedKeys,",")
    wscript.echo "  dc.RemoveAll": dc.RemoveAll
    wscript.echo "     dc.Count should be 0 -> " & dc.Count
    wscript.echo "----------- End DictCollection Demo of Basic Functionality ---------------------"
End Sub

' performs all available tests
Public Function SelfTest(DebugPrint)
    Dim retVal : retVal = True
    retVal = retVal And TestFunctionality(DebugPrint)
    retVal = retVal And TestCompatibility(DebugPrint)
    SelfTest = retVal
End Function

' combines errors with allErrors and prints errors to Immediate window if necessary
Private Function combineAndPrintTestErrors(testName, errors, allErrors, DebugPrint)
    Dim i 
    if UBound(errors) > -1 then wscript.echo "  FAIL: " & testName else wscript.echo "  OK: " & testName  
    For i = 0 to UBound(errors)
        UtilAddArrayValue allErrors, errors(i): If DebugPrint Then WscripT.echo "  " & errors(i)
    Next
End Function
' ======================== END: DEMO AND SELFTEST =================================================================================================

' ======================== START: FUNCTIONALITY TESTS =============================================================================================
' performs functionality test and returns true if passed and false if not
Public Function TestFunctionality(DebugPrint)
    TestFunctionality = UBound(TestFunctionalityErrors(DebugPrint)) > -1
End Function
' performs selftest and returns all errors  array
Public Function TestFunctionalityErrors(DebugPrint)
    Dim allErrors : allErrors = Array()
    wscript.echo "----------- Start DictCollection Functionality Test ----------------------------"
    combineAndPrintTestErrors "Basic Functionality", testBasicFunctionality(), allErrors, DebugPrint
    wscript.echo "----------- End DictCollection Functionality Test ------------------------------"
    TestFunctionalityErrors = allErrors
End Function

' Tests basic functionality that is also needed by Scripting.Dictionary and Collection emulation
' dc.CollectionType, dc.CompareMode
' dc.Add(val), dc.Add(val, key)
' dc.Item(key), dc.Item(index), dc(), dc(key), dc(index)
' dc.Remove(key), dc.Remove(index), dc.RemoveAll
Private Function testBasicFunctionality() 
    Dim errors, dc1, dc2, currentTest, i, val1, val2, key1, key2, missing
    errors = Array()
    missing = GetMissingValue(,1)
    
    Set dc1 = New DictCollection ' initialize class
currentTest = "[IOP-1] Initialized object properties - New DictCollection must have .Count=0": If dc1.Count <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest)
    currentTest = "[IOP-2] Initialized object properties - New DictCollection must have .CollectionType=1 (empty array)": If dc1.CollectionType <> 1 Then UtilAddArrayValue errors, ("-> " & currentTest)
    currentTest = "[IOP-3] Initialized object properties - Initial CompareMode must be binary": If dc1.CompareMode <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest)
    currentTest = "[IOP-4] Initialized object properties - Throw Errors must be fale by default": If dc1.ThrowErrors <> False Then UtilAddArrayValue errors, ("-> " & currentTest)
    currentTest = "[SAG-1] Simple add and get - Adding a single item": dc1.Add "key1", "item1"
    currentTest = "[SAG-2] Simple add and get - Retrieving single item": If dc1.Item("key1") <> "item1" Then UtilAddArrayValue errors, ("-> " & currentTest)
    currentTest = "[SAG-3] Simple add and get - Retrieving key for single item": If dc1.KeyOfItemAt(0) <> "key1" Then UtilAddArrayValue errors, ("-> " & currentTest)
    currentTest = "[SAG-4] Simple add and get - Retrieving index for key": If dc1.IndexOfKey("key1") <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest)
    
    Set dc1 = New DictCollection: dc1.Add, 100 ' initialize class and add value
currentTest = "[AOP-1] Array object properties - DictCollection with one item must have .Count=1": If dc1.Count <> 1 Then UtilAddArrayValue errors, ("-> " & currentTest)
    currentTest = "[AOP-2] Array object properties - DictCollection with one item and no keys must have .CollectionType=2 (filled array)": If dc1.CollectionType <> 2 Then UtilAddArrayValue errors, ("-> " & currentTest)
    
    Set dc1 = New DictCollection: dc1.Add "keyA", 100 ' initialize class and add value with key
currentTest = "[COP-1] Collection object properties - DictCollection with one item and one key must have .CollectionType=4 (filled key/value store)": If dc1.CollectionType <> 4 Then UtilAddArrayValue errors, ("-> " & currentTest)
    
    Set dc1 = New DictCollection: dc1.Add "keyA", 100: dc1.Add , 200 ' initialize class and add value with key and value without key
currentTest = "[COP-2] Collection object properties - DictCollection with one item and one key must have .CollectionType=5 (filled key/value store with items having no key)": If dc1.CollectionType <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
    
    Set dc1 = New DictCollection ' initialize class
currentTest = "[DV-1] Default values - Default NonExistingValue must be [NONEXISTING]": If dc1.NonExistingValue <> NONEXISTING_VALUE_DEFAULT Then UtilAddArrayValue errors, ("-> " & currentTest)
    currentTest = "[DV-2] Default values - Default EmptyCollectionValue must be [EMPTY]": If dc1.EmptyCollectionValue <> EMPTY_COLLECTION_VALUE_DEFAULT Then UtilAddArrayValue errors, ("-> " & currentTest)
    currentTest = "[DV-3] Default values - Empty DictCollection must have [EMPTY] as default value": If dc1 <> dc1.EmptyCollectionValue Then UtilAddArrayValue errors, ("-> " & currentTest)
    currentTest = "[DV-4] Default values - Nonexisting item must have [NONEXISTING] as default value": If dc1.item(0) <> dc1.NonExistingValue Then UtilAddArrayValue errors, ("-> " & currentTest)
    currentTest = "[DV-5] Default values - Nonexisting item must have [NONEXISTING] as default value": If dc1.item("nonexisting") <> dc1.NonExistingValue Then UtilAddArrayValue errors, ("-> " & currentTest)
    
currentTest = "[MDV-1] Modified default values for nonexisting items and empty collections"
    Set dc1 = New DictCollection: dc1.NonExistingValue = 1: dc1.EmptyCollectionValue = 2 ' initialize class and modify default values
    If dc1.NonExistingValue <> 1 Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.EmptyCollectionValue <> 2 Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1 <> 2 Or dc1 <> dc1.EmptyCollectionValue Then UtilAddArrayValue errors, ("-> " & currentTest)
    If CInt(dc1.item(0)) <> 1 Or dc1.item(0) <> dc1.NonExistingValue Then UtilAddArrayValue errors, ("-> " & currentTest)
    If CInt(dc1.item(1).item(2).item(3).item(4)) <> 1 Or dc1.item(1).item(2).item(3).item(4) <> dc1.NonExistingValue Then UtilAddArrayValue errors, ("-> " & currentTest)
    If CInt(dc1.item("nonexisting")) <> 1 Or dc1.item("nonexisting") <> dc1.NonExistingValue Then UtilAddArrayValue errors, ("-> " & currentTest)
    If CInt(dc1.item("a").item("b").item("c").item("d")) <> 1 Or dc1.item("a").item("b").item("c").item("d") <> dc1.NonExistingValue Then UtilAddArrayValue errors, ("-> " & currentTest)
    If CInt(dc1.item("")) <> 1 Or dc1.item("") <> dc1.NonExistingValue Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[MDV-2] Modified default values in nested DictCollections"
    dc1.Add "keyA", New DictCollection: dc1.Item(0).Add "keyB", New DictCollection
    ' VBScript Change: upwards propagation of default values does not work, maybe because of different default property evaluation
    'If CInt(dc1.Item(0)) <> 2 Or dc1.Item(0) <> dc1.EmptyCollectionValue Then UtilAddArrayValue errors, ("-> " & currentTest)
    'If CInt(dc1.Item("keyA")) <> 2 Or dc1.Item("keyA") <> dc1.EmptyCollectionValue Then UtilAddArrayValue errors, ("-> " & currentTest)    
    If CInt(dc1.Item(0).Item(0)) <> 2 Or dc1.Item(0).Item(0) <> dc1.EmptyCollectionValue Then UtilAddArrayValue errors, ("-> " & currentTest)
    If CInt(dc1.Item("keyA").Item("keyB")) <> 2 Or dc1.Item("keyA").Item("keyB") <> dc1.EmptyCollectionValue Then UtilAddArrayValue errors, ("-> " & currentTest)
    If CInt(dc1.Item(0).Item(0).Item(1)) <> 1 Or dc1.Item(0).Item(0).Item(1) <> dc1.NonExistingValue Then UtilAddArrayValue errors, ("-> " & currentTest)
    If CInt(dc1.Item("keyA").Item("keyB").Item("nonexisting")) <> 1 Or dc1.Item("keyA").Item("keyB").Item("nonexisting") <> dc1.NonExistingValue Then UtilAddArrayValue errors, ("-> " & currentTest)
    
    Set dc1 = New DictCollection
    Dim testArray(1) : testArray(0) = "TEST": testArray(1) = 222
    Set dc2 = New DictCollection: dc2.Add, "COL": dc2.Add, 333
    
currentTest = "[AIWAK-1] Adding items with ascending keys"
    dc1.RemoveAll: dc1.ThrowErrors = False: dc1.CompareMode = 0
    dc1.Add "keyA", "itemA"
    dc1.Add "keyB", 111
    dc1.Add "keyC", testArray
    dc1.Add "keyD", dc2
currentTest = "[AIWAK-2] Retrieving items with ascending keys by index":
    If dc1.Item(0) <> "itemA" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(1) <> 111 Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(2)(0) <> testArray(0) Or dc1.item(2)(1) <> testArray(1) Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(3).item(1) <> dc2.item(1) Or dc1.item(3).item(2) <> dc2.item(2) Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[AIWAK-3] Retrieving items with ascending keys by key"
    If dc1.Item("keyA") <> "itemA" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("keyB") <> 111 Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("keyC")(0) <> testArray(0) Or dc1.item("keyC")(1) <> testArray(1) Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("keyD").item(0) <> dc2.item(0) Or dc1.item("keyD").item(1) <> dc2.item(1) Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[AIWAK-4] Counting items with ascending keys": If dc1.Count <> 4 Then UtilAddArrayValue errors, ("-> " & currentTest)
    
currentTest = "[AIWDK-1] Adding items with descending keys"
    dc1.RemoveAll: dc1.ThrowErrors = False: dc1.CompareMode = 0
    dc1.Add "keyC", "itemC"
    dc1.Add "keyB", "itemB"
    dc1.Add "keyA", "itemA"
currentTest = "[AIWDK-2] Retrieving items with descending keys by index"
    If dc1.item(0) <> "itemC" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(1) <> "itemB" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(2) <> "itemA" Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[AIWDK-3] Retrieving items with descending keys by key"
    If dc1.item("keyA") <> "itemA" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("keyB") <> "itemB" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("keyC") <> "itemC" Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[AIWDK-4] Counting items with descending keys": If dc1.Count <> 3 Then UtilAddArrayValue errors, ("-> " & currentTest)
    
currentTest = "[AIWWK-1] Adding items with and without keys"
    dc1.RemoveAll: dc1.ThrowErrors = False: dc1.CompareMode = 0
    dc1.Add, "FirstArrayItem"
    dc1.Add "key1", "FirstKeyItem"
    dc1.Add "", "SecondArrayItem"
    dc1.Add "key2", "SecondKeyItem"
currentTest = "[AIWWK-2] Retrieving items items with and without keys by index"
    If dc1.item(0) <> "FirstArrayItem" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(1) <> "FirstKeyItem" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(2) <> "SecondArrayItem" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(3) <> "SecondKeyItem" Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[AIWWK-3] Retrieving items items with and without keys by index"
    If dc1.item("key1") <> "FirstKeyItem" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("key2") <> "SecondKeyItem" Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[AIWWK-4] Counting items with and without keys": If dc1.Count <> 4 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[AIWWK-5] Counting keys of items with and without keys": If dc1.KeyCount <> 2 Then UtilAddArrayValue errors, ("-> " & currentTest)
    
currentTest = "[AII-1] Accessing invalid items"
    dc1.RemoveAll: dc1.ThrowErrors = False: dc1.CompareMode = 0
    dc1.Add "keyC", "itemC"
    dc1.Add "keyB", "itemB"
    dc1.Add "keyA", "itemA"
currentTest = "[AII-2] Retrieving items with invalid index"
    If dc1.item(-1) <> dc1.NonExistingValue Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(3) <> dc1.NonExistingValue Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[AII-3] Retrieving items with invalid index"
    If dc1.item("") <> dc1.NonExistingValue Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("_") <> dc1.NonExistingValue Then UtilAddArrayValue errors, ("-> " & currentTest)
    dc1.NonExistingValue = "": val1 = dc1.item("keyc")
    If val1 <> "" Then UtilAddArrayValue errors, ("-> " & currentTest)
    'If dc1.item("keyc") <> "" Then UtilAddArrayValue errors, ("-> " & currentTest) ' <- This does not work. Left and right side of <> operator have to be either both properties or both values
    
currentTest = "[TCM-1] Testing binary CompareMode - Adding items"
    dc1.RemoveAll: dc1.ThrowErrors = False: dc1.CompareMode = 0
    dc1.Add "keya", "itema"
    dc1.Add "keyA", "itemA"
    dc1.Add "keyB", "itemB"
    dc1.Add "KEYC", "ITEMC"
    dc1.Add ChrW(8352), "EURO8352"
    dc1.Add Chr(128), "EURO128"
    dc1.Add Chr(164), "EURO164"
    dc1.Add ChrW(9702), "BULLET5" 'White Bullet
    dc1.Add ChrW(183), "BULLET1" 'Small Bullet
    dc1.Add ChrW(8226), "BULLET2" 'Fat Bullet
    dc1.Add ChrW(8729), "BULLET3" 'Bullet Operator
    dc1.Add ChrW(9679), "BULLET4" 'Black Cirlce
    dc1.Add "aàáâãäå", "LOWER_ASCII_A_CHARS" 'lower case variants of the ASCII character a
    dc1.Add "AÀÁÂÃÄÅ", "UPPER_ASCII_A_CHARS" 'upper case variants of the ASCII character a
    dc1.Add ChrW(257) & ChrW(259) & ChrW(261), "LOWER_UNICODE_A_CHARS"  'lower case variants of the UNICODE character a
    dc1.Add ChrW(256) & ChrW(258) & ChrW(260), "UPPER_UNICODE_A_CHARS"  'upper case variants of the UNICODE character a
    If dc1.Count <> 16 Or dc1.KeyCount <> 16 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[TCM-3] Retrieving items with case sensitive key comparison"
    If dc1.item("keya") <> "itema" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("keyA") <> "itemA" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("keyB") <> "itemB" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("KeYB") <> dc1.NonExistingValue Or dc1.item("KEYB") <> dc1.NonExistingValue Or dc1.item("keyb") <> dc1.NonExistingValue Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("KEYC") <> "ITEMC" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("KEyC") <> dc1.NonExistingValue Or dc1.item("Keyc") <> dc1.NonExistingValue Or dc1.item("keyc") <> dc1.NonExistingValue Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(Chr(128)) <> "EURO128" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(Chr(164)) <> "EURO164" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(ChrW(8352)) <> "EURO8352" Then UtilAddArrayValue errors, ("-> " & currentTest)
    'VBScript change: old line -> dc1.Add "€", "EURO" ' adding euro again with symbol
    dc1.Add Chr(128), "EURO" ' adding euro again with symbol
    If dc1.Count <> 16 Or dc1.KeyCount <> 16 Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(Chr(128)) <> "EURO" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(ChrW(183)) <> "BULLET1" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(ChrW(8226)) <> "BULLET2" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(ChrW(8729)) <> "BULLET3" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(ChrW(9679)) <> "BULLET4" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(ChrW(9702)) <> "BULLET5" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("aàáâãäå") <> "LOWER_ASCII_A_CHARS" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("AÀÁÂÃÄÅ") <> "UPPER_ASCII_A_CHARS" Then UtilAddArrayValue errors, ("-> " & currentTest) ' upper case keys will be sorted before lower case keys
    If dc1.item(ChrW(257) & ChrW(259) & ChrW(261)) <> "LOWER_UNICODE_A_CHARS" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(ChrW(256) & ChrW(258) & ChrW(260)) <> "UPPER_UNICODE_A_CHARS" Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[TCM-4] Change of CompareMode must leave keys intact"
    dc1.CompareMode = 1 ' case insensitive, converts unicode characters
    ' in VBScript, some Unicode Code Points are sorted differently than in VBA:
    ' | Index | VBA Key                         | VBScript Key 
    ' +-------+---------------------------------+------------------------------------
    ' | 0 - 7 | ... same order ...              | ... same order ...
    ' |  008  | "āăą" (LOWER_UNICODE_A_CHARS)   | "aàáâãäå" (LOWER_ASCII_A_CHARS)
    ' |  009  | "ĀĂĄ" (UPPER_UNICODE_A_CHARS)   | "AÀÁÂÃÄÅ" (UPPER_ASCII_A_CHARS)
    ' |  010  | "AÀÁÂÃÄÅ" (UPPER_ASCII_A_CHARS) | "ĀĂĄ" (UPPER_UNICODE_A_CHARS)
    ' |  011  | "aàáâãäå" (LOWER_ASCII_A_CHARS) | "āăą" (LOWER_UNICODE_A_CHARS)
    ' |  012  | "keyA"                          | "keyA"
    ' |  013  | "keya"                          | "keya"
    ' |  014  | "keyB"                          | "keyB"
    ' |  015  | "keyC"                          | "keyC"
    ' Also, in VBScript some Code Points are not equal when comparing them in case insensitive "Text Mode"
    ' VBA:       StrComp("AÀÁÂÃÄÅ","aàáâãäå",1) ' -> 0 = equal
    ' VBScript:  StrComp("AÀÁÂÃÄÅ","aàáâãäå",1) ' -> 1 = not equal    
    If dc1.item("keya") <> "itemA" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("keyA") <> "itemA" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("AÀÁÂÃÄÅ") <> "UPPER_ASCII_A_CHARS" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("aàáâãäå") <> "LOWER_ASCII_A_CHARS" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(ChrW(257) & ChrW(259) & ChrW(261)) <> "UPPER_UNICODE_A_CHARS" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(ChrW(256) & ChrW(258) & ChrW(260)) <> "UPPER_UNICODE_A_CHARS" Then UtilAddArrayValue errors, ("-> " & currentTest)
    dc1.CompareMode = 0 ' case sensitive again
    If dc1.item("keya") <> "itema" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("keyA") <> "itemA" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("aàáâãäå") <> "LOWER_ASCII_A_CHARS" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("AÀÁÂÃÄÅ") <> "UPPER_ASCII_A_CHARS" Then UtilAddArrayValue errors, ("-> " & currentTest) ' upper case keys will be sorted before lower case keys
    If dc1.item(ChrW(257) & ChrW(259) & ChrW(261)) <> "LOWER_UNICODE_A_CHARS" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(ChrW(256) & ChrW(258) & ChrW(260)) <> "UPPER_UNICODE_A_CHARS" Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[TCM-5] Change of CompareMode with empty key array must work"
    dc1.RemoveAll
    dc1.CompareMode = 0: If dc1.CompareMode <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest)
    dc1.CompareMode = 1: If dc1.CompareMode <> 1 Then UtilAddArrayValue errors, ("-> " & currentTest)
    dc1.CompareMode = 2: If dc1.CompareMode <> 2 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[TCM-6] Testing TextCompare Mode - Adding items with case insensitive keys"
    dc1.RemoveAll
    dc1.CompareMode = 1
    dc1.Add "keya", "itema"
    dc1.Add "KEYA", "ITEMA" 'overwrites item at 'keya'
    dc1.Add "keyB", "itemB"
    dc1.Add "KEYC", "ITEMC"
    dc1.Add "keyb", "itemb" 'overwrites item at 'keyb'
    dc1.Add ChrW(8352), "EURO8352"
    dc1.Add Chr(128), "EURO128"
    dc1.Add Chr(164), "EURO164"
    dc1.Add "KeyC", "ItemC" 'overwrites item at 'keyc'
    dc1.Add ChrW(9702), "BULLET5" 'White Bullet
    dc1.Add ChrW(183), "BULLET1" 'Small Bullet
    dc1.Add ChrW(8226), "BULLET2" 'Fat Bullet
    dc1.Add ChrW(8729), "BULLET3" 'Bullet Operator
    dc1.Add ChrW(9679), "BULLET4" 'Black Cirlce
    currentTest = "[TCM-6] Retrieving items with case insensitive key comparison"
    If dc1.item("keya") <> "ITEMA" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("KEYA") <> "ITEMA" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(0) <> "ITEMA" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("KeYb") <> "itemb" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(1) <> "itemb" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("keyc") <> "ItemC" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(2) <> "ItemC" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(Chr(128)) <> "EURO128" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(Chr(164)) <> "EURO164" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(ChrW(8352)) <> "EURO8352" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(Chr(128)) <> "EURO128" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(ChrW(183)) <> "BULLET1" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(ChrW(8226)) <> "BULLET2" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(ChrW(8729)) <> "BULLET3" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(ChrW(9679)) <> "BULLET4" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(ChrW(9702)) <> "BULLET5" Then UtilAddArrayValue errors, ("-> " & currentTest)

currentTest = "[AOA-1] Array only access"
    dc1.RemoveAll: dc1.ThrowErrors = False: dc1.CompareMode = 0
    dc1.Add, "ItemA"
    dc1.Add, 1.23
    dc1.Add, testArray
    dc1.Add, "4"
    dc1.Add, 5.67
currentTest = "[AOA-2] Testing array item count"
    If dc1.Count <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[AOA-3] Removing array items"
    dc1.Remove 1
   If dc1.Count <> 4 Then UtilAddArrayValue errors, ("-> " & currentTest)
    If (dc1.item(0) <> "ItemA") Or (dc1.item(1)(0) <> "TEST") Or (dc1.item(2) <> "4") Or (dc1.item(3) <> 5.67) Then UtilAddArrayValue errors, ("-> " & currentTest)
    dc1.Remove 3
    If dc1.Count <> 3 Then UtilAddArrayValue errors, ("-> " & currentTest)
    If (dc1.item(0) <> "ItemA") Or (dc1.item(1)(0) <> "TEST") Or (dc1.item(2) <> "4") Then UtilAddArrayValue errors, ("-> " & currentTest)
    dc1.Remove 0.9 'should remove item at index 1 because of Collection emulated behavior
    If (dc1.item(0) <> "ItemA") Or (dc1.item(1) <> "4") Then UtilAddArrayValue errors, ("-> " & currentTest)
    dc1.Remove Empty 'should not remove anything
    dc1.Remove "" 'should not remove anything
    dc1.Remove 0
    If dc1.Count <> 1 Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(0) <> "4" Then UtilAddArrayValue errors, ("-> " & currentTest)
    dc1.Remove dc1.Count - 1
    If dc1.Count <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest)
    dc1.Remove 0 ' should not throw errors
currentTest = "[AOA-4] Adding 1000 array items"
    For i =0 to  1000
        dc1.Add, i * 10
    Next
currentTest = "[AOA-5] Testing access of 1000 array items"
    For i =0 to 999 Step 3
        If dc1.item(i) <> i * 10 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Next
currentTest = "[AOA-6] Getting and setting array items"
    ' removing first 10 items and changin next 10 items to "X"
    For i = 0 to 9
        dc1.Remove 0: dc1.item(9) = "X"
    Next
    For i = 0 to 9
        If dc1.item(i) <> "X" Then UtilAddArrayValue errors, ("-> " & currentTest)
    Next
    
currentTest = "[MADU-1] Mixed Array and Dictionary usage"
    dc1.RemoveAll: dc1.ThrowErrors = False: dc1.CompareMode = 0
    dc1.Add, "ItemAWithoutKey"         '0 -> will be removed by index
    dc1.Add "Key1", "ItemBWithKey1"     '1 -> will be at index=0, item="ItemFWithKey1"
    dc1.Add "", testArray               '2 -> will be at index=1, item=testArray
    dc1.Add "Key2", "ItemCWithKey2"     '3 -> will be removed by index
    dc1.Add "key3", "ItemDWithKey3"     '4 -> will be at index=2, item=Empty
    dc1.Add, "ItemEWithoutKey"         '5 -> will be at index=3, item="ItemEWithoutKey"
    dc1.Add, Empty                     '6 -> will be at index=4, item="Hello"
    dc1.Add "789", testArray            '7 -> will be at index=5, item=testArray
    dc1.Add "456", "Hello"              '8 -> will be at index=6
    dc1.Add "123", "World"              '9 -> will be removed
    dc1.item(1) = "ItemFWithKey1"
    dc1.Remove 3
    dc1.item("key3") = dc1.item("Key2") ' should change to empty string
    dc1.Remove (dc1.Count - 1) ' remove last element
    dc1.item(5) = dc1.item("456")
    dc1.Remove 0
    If dc1.Count <> 7 Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item("") <> dc1.NonExistingValue Then UtilAddArrayValue errors, ("-> " & currentTest)
    If (dc1.item(0) <> "ItemFWithKey1") Or (dc1.item("Key1") <> "ItemFWithKey1") Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(1)(1) <> 222 Then UtilAddArrayValue errors, ("-> " & currentTest)
    If (dc1.item("Key2") <> dc1.NonExistingValue) Or dc1.Exists("Key2") Then UtilAddArrayValue errors, ("-> " & currentTest)
    If (dc1.item(2) <> dc1.NonExistingValue) Or (Not dc1.Exists("key3")) Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(3) <> "ItemEWithoutKey" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If (dc1.item(4) <> dc1.item("456")) Or (dc1.item("456") <> "Hello") Then UtilAddArrayValue errors, ("-> " & currentTest)
    If (dc1.item(5)(0) <> "TEST") Or (dc1.item("789")(1) <> 222) Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(6) <> "Hello" Then UtilAddArrayValue errors, ("-> " & currentTest)
    
    ' remove 3 items at the beginning -> itemCount=4, keyCount = 2
    dc1.Remove 0: dc1.Remove 0: dc1.Remove 0:
    If (dc1.Count <> 4) Or (dc1.KeyCount <> 2) Then UtilAddArrayValue errors, ("-> " & currentTest)
    
    ' remove 2 items at the end -> itemCount=2, keyCount=0
    dc1.Remove (dc1.Count - 1): dc1.Remove (dc1.Count - 1)
    If (dc1.Count <> 2) Or (dc1.KeyCount <> 0) Then UtilAddArrayValue errors, ("-> " & currentTest)
    
    ' remove remaining 2 items
    dc1.Remove (dc1.Count - 1): dc1.Remove 0
    If (dc1.Count <> 0) Or (dc1.KeyCount <> 0) Then UtilAddArrayValue errors, ("-> " & currentTest)
    
currentTest = "[MODK-1] Manipulation of Dictionary keys - Replacing existing keys with existing keys"
    dc1.RemoveAll: dc1.ThrowErrors = False: dc1.CompareMode = 0
    dc1.Add "Key1", "Item1"
    dc1.Add "Key2", "Item2"
    dc1.Add "Key3", "Item3"
    dc1.Add "Key4", "Item4"
    dc1.Add "Key5", "Item5"
    dc1.Key("Key1") = "Key2" ' -> "Item2" should be dropped
    dc1.Key("Key5") = "Key4" ' -> "Item4" should be dropped
    dc1.Key("Key2") = "Key1" ' -> "Key1" should again point to "Item1"
    dc1.Key("Key4") = "Key5" ' -> "Key5" should again point to "Item5"
    If (dc1.Count <> 3) Or (dc1.KeyCount <> 3) Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(0) <> "Item1" Or dc1.item("Key1") <> "Item1" Or dc1.KeyOfItemAt(0) <> "Key1" Or dc1.IndexOfKey("Key1") <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(1) <> "Item3" Or dc1.item("Key3") <> "Item3" Or dc1.KeyOfItemAt(1) <> "Key3" Or dc1.IndexOfKey("Key3") <> 1 Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(2) <> "Item5" Or dc1.item("Key5") <> "Item5" Or dc1.KeyOfItemAt(2) <> "Key5" Or dc1.IndexOfKey("Key5") <> 2 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[MODK-2] Replacing existing and nonexisting keys with nonexisting keys"
    dc1.Key("Key1") = "xyz"
    dc1.Key("Key2") = "efg" ' should do nothing because "Key2" does not exist anymore
    dc1.Key("Key3") = "abc"
    dc1.Key("Key5") = "abcd"
    If dc1.item(0) <> "Item1" Or dc1.item("xyz") <> "Item1" Or dc1.KeyOfItemAt(0) <> "xyz" Or dc1.IndexOfKey("xyz") <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(1) <> "Item3" Or dc1.item("abc") <> "Item3" Or dc1.KeyOfItemAt(1) <> "abc" Or dc1.IndexOfKey("abc") <> 1 Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(2) <> "Item5" Or dc1.item("abcd") <> "Item5" Or dc1.KeyOfItemAt(2) <> "abcd" Or dc1.IndexOfKey("abcd") <> 2 Then UtilAddArrayValue errors, ("-> " & currentTest)
    If (dc1.Count <> 3) Or (dc1.KeyCount <> 3) Then UtilAddArrayValue errors, ("-> " & currentTest)
    dc1.Key("abcd") = ""
    If (dc1.Count <> 3) Or (dc1.KeyCount <> 2) Or dc1.item("abcd") <> Empty Then UtilAddArrayValue errors, ("-> " & currentTest)
    dc1.Key("abc") = ""
    dc1.Key("xyz") = Empty
    If (dc1.Count <> 3) Or (dc1.KeyCount <> 0) Then UtilAddArrayValue errors, ("-> " & currentTest)
    
currentTest = "[ISRII-1] Inserting, Setting and Removing Items by Index - Inserting Items by Index"
    dc1.RemoveAll: dc1.ThrowErrors = False: dc1.CompareMode = 0
    dc1.Insert "e", 0, missing   ' -> inserted at 0
    dc1.Insert "c", 0, missing   ' -> inserted before "e" at 0, "e" is now at 1
    dc1.Insert "d", 1, missing   ' -> inserted before "e" at 1, "e" is not at 2
    dc1.Insert "f", 3, missing   ' -> inserted after "e" at 3
    dc1.Insert "b", 0, missing   ' -> inserted at 0 before "c", all others move one up
    dc1.Insert "a", 0, missing   ' -> inserted at 0 before "b", all others move one up
    dc1.Insert "h", 7, missing   ' -> inserted at 7, creates empty item at 6
    If dc1.item(0) <> "a" Or dc1.item(1) <> "b" Or dc1.item(2) <> "c" Or dc1.item(3) <> "d" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.item(4) <> "e" Or dc1.item(5) <> "f" Or dc1.item(6) <> Empty Or dc1.item(7) <> "h" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.Count <> 8 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[ISRII-2] Adding and Removing Items by Index - Setting Items by Index"
    dc1.RemoveAll: dc1.ThrowErrors = False: dc1.CompareMode = 0
    dc1.item(3) = "d": dc1.item(1) = "b": dc1.item(0) = "a": dc1.item(4) = "e"
    If dc1.item(0) <> "a" Or dc1.item(1) <> "b" Or dc1.item(2) <> Empty Or dc1.item(3) <> "d" Or dc1.item(4) <> "e" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If dc1.Count <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[ISRII-3] Adding and Removing Items by Index - Removing Items by Index"
    dc1.Remove 4: If dc1.item(3) <> "d" Or dc1.item(4) <> dc1.NonExistingValue Or dc1.Count <> 4 Then UtilAddArrayValue errors, ("-> " & currentTest)
    dc1.Remove 2: If dc1.item(1) <> "b" Or dc1.item(2) <> "d" Or dc1.item(3) <> dc1.NonExistingValue Or dc1.Count <> 3 Then UtilAddArrayValue errors, ("-> " & currentTest)
    dc1.Remove 0: If dc1.item(0) <> "b" Or dc1.item(1) <> "d" Or dc1.item(2) <> dc1.NonExistingValue Or dc1.Count <> 2 Then UtilAddArrayValue errors, ("-> " & currentTest)
    dc1.Remove 0: If dc1.item(0) <> "d" Or dc1.item(1) <> dc1.NonExistingValue Or dc1.Count <> 1 Then UtilAddArrayValue errors, ("-> " & currentTest)
    dc1.Remove 0: If dc1.item(0) <> dc1.NonExistingValue Or dc1.Count <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest)
    dc1.Remove 0: If dc1.item(0) <> dc1.NonExistingValue Or dc1.Count <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest)
    
    For i = 1 To 2
        If i = 1 Then
currentTest = "[AMSK-1] Adding multiple items with same keys with eager sorting": Set dc1 = New DictCollection: dc1.LazySorting = False
        Else
currentTest = "[AMSK-2] Adding multiple items with same keys with lazy sorting": Set dc1 = New DictCollection: dc1.LazySorting = True
        End If
        dc1.Add "b", 222.1: dc1.Add "d", 444.1: dc1.Add "c", 333.1: dc1.Add "c", 333.2
        dc1.Add "c", 333 ' final index = 2
        dc1.Add "b", 222.2: dc1.Add "d", 444.2
        dc1.Add "b", 222 ' final index = 0
        dc1.Add "a", 111.1
        dc1.Add "d", 444 ' final index = 1
        dc1.Add "a", 111.2
        dc1.Add "e", 555 ' final index = 4
        dc1.Add "f", 666.1: dc1.Add "a", 111.3
        dc1.Add "a", 111 ' final index = 3
        dc1.Add "f", 666.1
        dc1.Add "f", 666 ' final index = 5
        If dc1.item(3) <> 111 Then UtilAddArrayValue errors, ("-> " & currentTest & " dc(3) must equal 111")
        If dc1.item(0) <> 222 Then UtilAddArrayValue errors, ("-> " & currentTest & " dc(0) must equal 222")
        If dc1.item(2) <> 333 Then UtilAddArrayValue errors, ("-> " & currentTest & " dc(2) must equal 333")
        If dc1.item(1) <> 444 Then UtilAddArrayValue errors, ("-> " & currentTest & " dc(1) must equal 444")
        If dc1.item(4) <> 555 Then UtilAddArrayValue errors, ("-> " & currentTest & " dc(4) must equal 555")
        If dc1.item(5) <> 666 Then UtilAddArrayValue errors, ("-> " & currentTest & " dc(5) must equal 666")
        If dc1.item(6) <> dc1.NonExistingValue Then UtilAddArrayValue errors, ("-> " & currentTest & " dc(6) must equal " & dc1.NonExistingValue)
        If dc1.item("a") <> 111 Then UtilAddArrayValue errors, ("-> " & currentTest & " dc(""a"") must equal 111")
        If dc1.item("b") <> 222 Then UtilAddArrayValue errors, ("-> " & currentTest & " dc(""b"") must equal 222")
        If dc1.item("c") <> 333 Then UtilAddArrayValue errors, ("-> " & currentTest & " dc(""c"") must equal 333")
        If dc1.item("d") <> 444 Then UtilAddArrayValue errors, ("-> " & currentTest & " dc(""d"") must equal 444")
        If dc1.item("e") <> 555 Then UtilAddArrayValue errors, ("-> " & currentTest & " dc(""e"") must equal 555")
        If dc1.item("f") <> 666 Then UtilAddArrayValue errors, ("-> " & currentTest & " dc(""f"") must equal 666")
        If dc1.item("x") <> dc1.NonExistingValue Then UtilAddArrayValue errors, ("-> " & currentTest & " dc(""x"") must equal " & dc1.NonExistingValue)

        If i = 1 Then
currentTest = "[CK-1] Changing keys with eager sorting": Set dc1 = New DictCollection: dc1.LazySorting = False
        Else
currentTest = "[CK-2] Changing keys with lazy sorting": Set dc1 = New DictCollection: dc1.LazySorting = True
        End If

        dc1.RemoveAll: dc1.Add "keyA", "item1": dc1.Add "keyC", "item2": dc1.Add "keyE", "item3": dc1.Add "keyF", "item4"
        key1 = "keyA": val1 = "item1": If dc1.item(key1) <> val1 Then UtilAddArrayValue errors, ("-> " & currentTest & " dc(""" & key1 & """) must equal """ & val1 & """")
        key1 = "keyC": val1 = "item2": If dc1.item(key1) <> val1 Then UtilAddArrayValue errors, ("-> " & currentTest & " dc(""" & key1 & """) must equal """ & val1 & """")
        ' change keys to unused keys
        dc1.Key("keyA") = "keyB": key1 = "keyB": val1 = "item1": If dc1.item(key1) <> val1 Then UtilAddArrayValue errors, ("-> " & currentTest & " dc(""" & key1 & """) must equal """ & val1 & """ after change")
        dc1.Key("keyC") = "keyD": key1 = "keyD": val1 = "item2": If dc1.item(key1) <> val1 Then UtilAddArrayValue errors, ("-> " & currentTest & " dc(""" & key1 & """) must equal """ & val1 & """ after change")
        dc1.Key("keyF") = "keyA": key1 = "keyA": val1 = "item4": If dc1.item(key1) <> val1 Then UtilAddArrayValue errors, ("-> " & currentTest & " dc(""" & key1 & """) must equal """ & val1 & """ after change")
        dc1.Key("keyE") = "keyF": key1 = "keyF": val1 = "item3": If dc1.item(key1) <> val1 Then UtilAddArrayValue errors, ("-> " & currentTest & " dc(""" & key1 & """) must equal """ & val1 & """ after change")
        dc1.Key("keyB") = "keyG": key1 = "keyG": val1 = "item1": If dc1.item(key1) <> val1 Then UtilAddArrayValue errors, ("-> " & currentTest & " dc(""" & key1 & """) must equal """ & val1 & """ after change")
        If dc1.item("keyG") <> "item1" Or dc1.item("keyD") <> "item2" Or dc1.item("keyA") <> "item4" Or dc1.item("keyF") <> "item3" Then UtilAddArrayValue errors, ("-> wrong key/item associations after change.")
        ' change keys to used keys -> drop items that used those keys before
        ' item1 will get "keyA" and item4 will be dropped, still existing after that:  "keyA"=item1, "keyD"=item2, "keyF"=item3
        dc1.Key("keyG") = "keyA": key1 = "keyA": val1 = "item1": If dc1.item(key1) <> val1 Or dc1.Exists("keyG") Or (dc1.Count <> 3) Or (dc1.KeyCount <> 3) Then UtilAddArrayValue errors, ("-> " & currentTest & " item4 should not exist after assigning its key to item1")
        ' item3 will get "keyD" and item2 will be dropped, still existing after that: "keyA"=item1, "keyD"="item3
        dc1.Key("keyF") = "keyD": key1 = "keyD": val1 = "item3": If dc1.item(key1) <> val1 Or dc1.Exists("keyF") Or (dc1.Count <> 2) Or (dc1.KeyCount <> 2) Then UtilAddArrayValue errors, ("-> " & currentTest & " item2 should not exist after assigning its key to item1")
        ' item3 will get "keyA" and item1 will be dropped, still existing after that: "keyD"="item3
        dc1.Key("keyD") = "keyA": key1 = "keyA": val1 = "item3": If dc1.item(key1) <> val1 Or dc1.Exists("keyD") Or (dc1.Count <> 1) Or (dc1.KeyCount <> 1) Then UtilAddArrayValue errors, ("-> " & currentTest & " item1 should not exist after assigning its key to item1")
        ' item3 key will be set to none, still existing after that: item3 without key
        dc1.Key("keyA") = "": If dc1.item(0) <> "item3" Or (dc1.KeyCount <> 0) Then UtilAddArrayValue errors, ("-> " & currentTest & ": all keys must be removed after dropping all keys by setting them")
    
    
    Next
    testBasicFunctionality = errors
End Function
' ======================== END: FUNCTIONALITY TESTS =================================================================================================



' ======================== START: COMPATIBILITY TESTS ===============================================================================================
Public Function TestCompatibility(DebugPrint): TestCompatibility = UBound(TestCompatibilityErrors(DebugPrint)) > -1: End Function

Public Function TestCompatibilityErrors(DebugPrint)
    Dim allErrors: allErrors = Array()
    Wscript.echo "----------- Start DictCollection Compatibility Test ----------------------------"
    
    combineAndPrintTestErrors "Scripting Dictionary Compatibility using Dictionary", testScriptingDictionaryCompatibility(CreateObject("Scripting.Dictionary")), allErrors, DebugPrint
    combineAndPrintTestErrors "Scripting Dictionary Compatibility using DictCollection", testScriptingDictionaryCompatibility(New DictCollection), allErrors, DebugPrint
    
    combineAndPrintTestErrors "Collection Compatibility using DictCollection", testCollectionCompatibility(New DictCollection), allErrors, DebugPrint
    set d = new DictCollection: d.EmulateCollection=true
    Wscript.echo "----------- End DictCollection Compatibility Test ------------------------------"
    TestCompatibilityErrors = allErrors
End Function

Private Function testScriptingDictionaryCompatibility(d)
    Dim errors, currentTest, i, j, r, removeCount, indexToBeRemoved, indexToBeSwapped1, indexToBeSwapped2
    Dim test, Items, Keys, testdata, val1, val2
    Dim o1, o2, o3, o4, o5, o6
    errors = Array(): Items = Array(): Keys = Array()
    
    If TypeName(d) = "DictCollection" Then d.EmulateDictionary = True ' switch emulation on if running with DictCollection

currentTest = "[DES-1] Dictionary Empty State - Count and CompareMode must be zero"
    If d.Count <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.CompareMode <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[DES-2] Dictionary Empty State - .Items and .Keys must return initialized empty variant array"
    Items = d.Items: Keys = d.Keys
    If TypeName(Items) <> "Variant()" Or TypeName(Keys) <> "Variant()" Then
        UtilAddArrayValue errors, ("-> " & currentTest)
    Else
        If UBound(Items) <> -1 Or UBound(Keys) <> -1 Then UtilAddArrayValue errors, ("-> " & currentTest)
    End If
    For test = 1 To 9
        Select Case test
            Case 1: ' String
                currentTest = "[DBF-1] Dictionary Basic Functions - Adding, accessing, changing and removing items with keys of datatype String"
                testdata = Array(Array("keya", "a"), Array("keyb", "b"), Array("keyc", "c"), Array("keyd", "d"), Array("keye", "e"), Array("", "f")) ' Empty is tested in Case 10
            Case 2: ' Integer
                currentTest = "[DBF-2] Dictionary Basic Functions - Adding, accessing, changing and removing items with keys of datatype Integer"
                testdata = Array(Array(CInt(-10000), "a"), Array(CInt(-1), "b"), Array(CInt(0), "c"), Array(CInt(1), "d"), Array(CInt(10000), "e"), Array(CInt(32767), "f"))
            Case 3: ' Long
                currentTest = "[DBF-3] Dictionary Basic Functions - Adding, accessing, changing and removing items with keys of datatype Long"
                testdata = Array(Array(CLng(-100000000), "a"), Array(CLng(-1), "b"), Array(CInt(0), "c"), Array(CLng(1), "d"), Array(CLng(100000000), "e"), Array(CLng(2147483647), "f"))
            Case 4: ' Single
                currentTest = "[DBF-4] Dictionary Basic Functions - Adding, accessing, changing and removing items with keys of datatype Single"
                testdata = Array(Array(CSng(-1000.001), "a"), Array(CSng(-1.000001), "b"), Array(CSng(0), "c"), Array(CSng(1000.001), "d"), Array(CSng(1.000001), "e"), Array(CSng(3.402823E+38), "f"))
            Case 5: ' Double
                currentTest = "[DBF-5] Dictionary Basic Functions - Adding, accessing, changing and removing items with keys of datatype Double"
                testdata = Array(Array(CDbl(-100000.000000001), "a"), Array(CDbl(-1.00000000000001), "b"), Array(CDbl(0), "c"), Array(CDbl(100000.000000001), "d"), Array(CDbl(1.00000000000001), "e"), Array(CDbl(4.94065645841247E-324), "f"))
            Case 6: ' Currency
                currentTest = "[DBF-6] Dictionary Basic Functions - Adding, accessing, changing and removing items with keys of datatype Currency"
                testdata = Array(Array(CCur(-100000000000.0001), "a"), Array(CCur(-1.0001), "b"), Array(CCur(0), "c"), Array(CCur(1.0001), "d"), Array(CCur(100000000000.00019), "e"), Array(CCur(922337203685477.5), "f"))
            Case 7: ' Date
                currentTest = "[DBF-8] Dictionary Basic Functions - Adding, accessing, changing and removing items with keys of datatype Date"
                testdata = Array(Array(#1/1/100#, "a"), Array(#1/1/1899 10:00:01 AM#, "b"), Array(#1/1/1900#, "c"), Array(#12/31/2000 12:00:00 PM#, "d"), Array(#5/5/2100 11:00:00 PM#, "e"), Array(CDate(#12/31/9999 11:59:59 PM#), "f"))
            Case 8: ' Object
                currentTest = "[DBF-9] Dictionary Basic Functions - Adding, accessing, changing and removing items with keys of datatype Object"
                Set o1 = New DictCollection: Set o2 = New DictCollection: Set o3 = New DictCollection: Set o4 = New DictCollection: Set o5 = New DictCollection: Set o6 = New DictCollection
                testdata = Array(Array(o1, "a"), Array(o2, "b"), Array(o3, "c"), Array(o4, "d"), Array(o5, "e"), Array(o6, "f"))
            Case 9: ' Mixed data types
                currentTest = "[DBF-11] Dictionary Basic Functions - Adding, accessing, changing and removing items with mixed key datatypes"
                Set o1 = New DictCollection: Set o2 = New DictCollection: o1.DefaultValueEnabled = False: o2.DefaultValueEnabled = False
                testdata = Array(Array(Empty, "empty1"), Array(o1, "object1"), Array(CInt(1), "Number=1"), Array("1", "String='1'"), Array(1.00001, "Number=1.00001"), Array(-1.239847E-10, "Number=-1.239847E-10"), Array(o2, "object2"), Array(#12/31/2012 11:12:00 PM#, "Date='12/31/2012 11:12:00 PM'"))
        End Select
        d.RemoveAll
        For i = 0 To UBound(testdata): d.Add testdata(i)(0), testdata(i)(1): Next

        removeCount = UBound(testdata)
        For r = 0 To removeCount
            Items = d.Items: Keys = d.Keys
            If d.Count <> UBound(testdata) + 1 Then UtilAddArrayValue errors, ("-> " & currentTest & ": .Count must equal " & UBound(testdata) + 1)
            For i = 0 To UBound(testdata)
                ' compare type of .Item
                val1 = TypeName(d.Item(testdata(i)(0))): val2 = TypeName(testdata(i)(1))
                If val1 <> val2 Then UtilAddArrayValue errors, ("-> " & currentTest & ": .Item(" & testdata(i)(0) & ") is " & val1 & " but must be of type " & val2)
                ' compare value of .Item
                If d.Item(testdata(i)(0)) <> testdata(i)(1) Then UtilAddArrayValue errors, ("-> " & currentTest & ": .Item(" & testdata(i)(0) & ") must be " & testdata(i)(1))
                ' test .Exists()
                If d.Exists(testdata(i)(0)) <> True Then UtilAddArrayValue errors, ("-> " & currentTest & ": .Exists(" & testdata(i)(0) & ") must be true")
                ' compare type of item in .Items array
                val1 = TypeName(Items(i)): val2 = TypeName(testdata(i)(1))
                If val1 <> val2 Then UtilAddArrayValue errors, ("-> " & currentTest & ": .Items(" & i & ") is " & val1 & " but must be of type " & val2)
                ' compare value of item in .Items array
                If Items(i) <> testdata(i)(1) Then UtilAddArrayValue errors, ("-> " & currentTest & ": .Items(" & i & ") must be " & testdata(i)(1))
                ' compare type of key in .Keys array
                val1 = TypeName(Keys(i)): val2 = TypeName(testdata(i)(0))
                If val1 <> val2 Then UtilAddArrayValue errors, ("-> " & currentTest & ": .Keys(" & i & ") is " & val1 & " but must be of type " & val2)
                ' compare value of key in .Keys array
                If IsObject(Keys(i)) And IsObject(testdata(i)(0)) Then
                    If Not (Keys(i) Is testdata(i)(0)) Then UtilAddArrayValue errors, ("-> " & currentTest & ": .Keys(" & i & ") must be identical object as original key")
                Else
                    If Keys(i) <> testdata(i)(0) Then UtilAddArrayValue errors, ("-> " & currentTest & ": .Keys(" & i & ") must be " & testdata(i)(0))
                End If
            Next
            ' identify index to be removed, and indexes to be swapped
            indexToBeSwapped1 = -1: indexToBeSwapped1 = -1 ' do not swap anything
            Select Case r
                Case 0, 1:
                    indexToBeRemoved = UBound(testdata) \ 2 ' take item from the middle
                    If UBound(testdata) >= 3 Then indexToBeSwapped1 = indexToBeRemoved + 1: indexToBeSwapped2 = indexToBeRemoved - 1 ' swap values above and below index
                Case Else:
                    If UBound(testdata) Mod 2 = 0 Then
                        indexToBeRemoved = 0 ' take first index
                        If UBound(testdata) >= 2 Then indexToBeSwapped1 = 1: indexToBeSwapped2 = 2 ' swap others above
                    Else
                        indexToBeRemoved = UBound(testdata) ' take last
                        If UBound(testdata) >= 2 Then indexToBeSwapped1 = indexToBeRemoved - 1: indexToBeSwapped2 = indexToBeRemoved - 2 ' swap others below
                    End If
            End Select
            If indexToBeSwapped1 > -1 And indexToBeSwapped2 > -1 Then
                ' swap values in Dictionary
                If IsObject(d.Item(testdata(indexToBeSwapped1)(0))) Then Set d.Item(testdata(indexToBeRemoved)(0)) = d.Item(testdata(indexToBeSwapped1)(0)) Else d.Item(testdata(indexToBeRemoved)(0)) = d.Item(testdata(indexToBeSwapped1)(0))
                If IsObject(d.Item(testdata(indexToBeSwapped2)(0))) Then Set d.Item(testdata(indexToBeSwapped1)(0)) = d.Item(testdata(indexToBeSwapped2)(0)) Else d.Item(testdata(indexToBeSwapped1)(0)) = d.Item(testdata(indexToBeSwapped2)(0))
                If IsObject(d.Item(testdata(indexToBeRemoved)(0))) Then Set d.Item(testdata(indexToBeSwapped2)(0)) = d.Item(testdata(indexToBeRemoved)(0)) Else d.Item(testdata(indexToBeSwapped2)(0)) = d.Item(testdata(indexToBeRemoved)(0))
                ' swap values in test array
                If IsObject(testdata(indexToBeSwapped1)(1)) Then Set testdata(indexToBeRemoved)(1) = testdata(indexToBeSwapped1)(1) Else testdata(indexToBeRemoved)(1) = testdata(indexToBeSwapped1)(1)
                If IsObject(testdata(indexToBeSwapped2)(1)) Then Set testdata(indexToBeSwapped1)(1) = testdata(indexToBeSwapped2)(1) Else testdata(indexToBeSwapped1)(1) = testdata(indexToBeSwapped2)(1)
                If IsObject(testdata(indexToBeRemoved)(1)) Then Set testdata(indexToBeSwapped2)(1) = testdata(indexToBeRemoved)(1) Else testdata(indexToBeSwapped2)(1) = testdata(indexToBeRemoved)(1)
            End If
            d.Remove testdata(indexToBeRemoved)(0) ' remove 1 item in Dictionary
            UtilRemoveArrayValueByIndex testdata, indexToBeRemoved ' remove 1 item in test array
        Next
    Next

currentTest = "[DIKA-1] Dictionary Implicit Key Adding - Adding keys by accessing non-existing keys"
    d.RemoveAll
    val1 = d.Item("a")
    If val1 <> Empty Then UtilAddArrayValue errors, ("-> " & currentTest & ": nonexisting item must be empty")
    If Not d.Exists("a") Then UtilAddArrayValue errors, ("-> " & currentTest & ": nonexisting item must be empty")
    If d.Count <> 1 Then UtilAddArrayValue errors, ("-> " & currentTest & ": .Count must be different after acessing nonexisting key")
    val1 = d.Item(0)
    If val1 <> Empty Then UtilAddArrayValue errors, ("-> " & currentTest & ": nonexisting item must be empty")
    If Not d.Exists(0) Then UtilAddArrayValue errors, ("-> " & currentTest & ": nonexisting item must be empty")
    If d.Count <> 2 Then UtilAddArrayValue errors, ("-> " & currentTest & ": .Count must be different after acessing nonexisting key")
    Set o1 = New DictCollection: o1.DefaultValueEnabled = False
    val1 = d.Item(o1)
    If val1 <> Empty Then UtilAddArrayValue errors, ("-> " & currentTest & ": nonexisting item must be empty")
    If Not d.Exists(o1) Then UtilAddArrayValue errors, ("-> " & currentTest & ": nonexisting item must be empty")
    If d.Count <> 3 Then UtilAddArrayValue errors, ("-> " & currentTest & ": .Count must be different after acessing nonexisting key")
    Items = d.Items
    If Items(0) <> Empty Or Items(1) <> Empty Or Items(2) <> Empty Then UtilAddArrayValue errors, ("-> " & currentTest & ": implicitly added items must be empty!")
    Keys = d.Keys
    If TypeName(Keys(0)) <> "String" Or TypeName(Keys(1)) <> "Integer" Or Not IsObject(Keys(2)) Then
        UtilAddArrayValue errors, ("-> " & currentTest & ": implicitly added keys must have correct types!")
    Else
        If Keys(0) <> "a" Or Keys(1) <> 0 Or Not (Keys(2) Is o1) Then UtilAddArrayValue errors, ("-> " & currentTest & ": implicitly added keys must must have correct values!")
    End If
    
    
    
currentTest = "[DOAV-1] Dictionary Objects as Values - Adding, changing and retrieving objects as items"
    d.RemoveAll: Set o1 = New DictCollection: Set o2 = New DictCollection: Set o3 = New DictCollection: Set o4 = New DictCollection: Set o5 = New DictCollection: Set o6 = New DictCollection
    o1.DefaultValueEnabled = False: o2.DefaultValueEnabled = False: o3.DefaultValueEnabled = False: o4.DefaultValueEnabled = False: o5.DefaultValueEnabled = False: o6.DefaultValueEnabled = False
    d.Add "a", o1: d.Add "b", o2: d.Add "c", o3: d.Add "d", o4: d.Add "e", o5: d.Add "f", o6:
    If (Not IsObject(d.Item("a"))) Or (Not IsObject(d.Item("b"))) Or (Not IsObject(d.Item("c"))) Or (Not IsObject(d.Item("d"))) Or (Not IsObject(d.Item("e"))) Or (Not IsObject(d.Item("f"))) Then
        UtilAddArrayValue errors, ("-> " & currentTest & ": IsObject(.Item(...)) must return true")
    ElseIf (Not d.Item("a") Is o1) Or (Not d.Item("b") Is o2) Or (Not d.Item("c") Is o3) Or (Not d.Item("d") Is o4) Or (Not d.Item("e") Is o5) Or (Not d.Item("f") Is o6) Then
        UtilAddArrayValue errors, ("-> " & currentTest & ": obj Is .Item(key) must return true")
    End If
    ' swapping item "a" and "b", removing item "f"
    Set d.Item("f") = d.Item("a"): Set d.Item("a") = d.Item("b"):  Set d.Item("b") = d.Item("f"): d.Remove ("f")
    If (Not IsObject(d.Item("a"))) Or (Not IsObject(d.Item("b"))) Or (Not IsObject(d.Item("c"))) Or (Not IsObject(d.Item("d"))) Or (Not IsObject(d.Item("e"))) Then
        UtilAddArrayValue errors, ("-> " & currentTest & ": IsObject(.Item(...)) must return true after swapping")
    ElseIf (Not d.Item("a") Is o2) Or (Not d.Item("b") Is o1) Or (Not d.Item("c") Is o3) Or (Not d.Item("d") Is o4) Or (Not d.Item("e") Is o5) Then
        UtilAddArrayValue errors, ("-> " & currentTest & ": obj Is .Item(key) must return true after swapping")
    End If
    If d.Exists("f") Then UtilAddArrayValue errors, ("-> " & currentTest & ": key 'f' must not exist after swapping")
    If d.Count <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest & ": .Count must be 5 after swapping")
    ' removing item 0, 4, seting "d" to Integer
    d.Remove ("a"): d.Remove ("e"): d.Item("d") = 100
    If (Not IsObject(d.Item("b"))) Or (Not IsObject(d.Item("c"))) Or (Not IsNumeric(d.Item("d"))) Then
        UtilAddArrayValue errors, ("-> " & currentTest & ": wrong item datatypes after remove")
    ElseIf (Not d.Item("b") Is o1) Or (Not d.Item("c") Is o3) Or d.Item("d") <> 100 Then
        UtilAddArrayValue errors, ("-> " & currentTest & ": wrong item values after remove")
    End If
    If d.Exists("a") Or d.Exists("e") Then UtilAddArrayValue errors, ("-> " & currentTest & ": key 'a' and 'e' must not exist after swapping")
    If d.Count <> 3 Then UtilAddArrayValue errors, ("-> " & currentTest & ": .Count must be 3 after removing")
    
currentTest = "[DCM-1] Dictionary CompareMode - Adding items with binary compare mode"
    d.RemoveAll: d.CompareMode = 0
    d.Add "keya", "itema"
    d.Add "keyA", "itemA"
    d.Add "keyB", "itemB"
    d.Add "KEYC", "ITEMC"
    d.Add ChrW(8352), "EURO8352"
    d.Add Chr(128), "EURO128"
    d.Add Chr(164), "EURO164"
    d.Add ChrW(9702), "BULLET5" 'White Bullet
    d.Add ChrW(183), "BULLET1" 'Small Bullet
    d.Add ChrW(8226), "BULLET2" 'Fat Bullet
    d.Add ChrW(8729), "BULLET3" 'Bullet Operator
    d.Add ChrW(9679), "BULLET4" 'Black Cirlce
    d.Add "aàáâãäå", "LOWER_ASCII_A_CHARS" 'lower case variants of the ASCII character a
    d.Add "AÀÁÂÃÄÅ", "UPPER_ASCII_A_CHARS" 'upper case variants of the ASCII character a
    d.Add ChrW(257) & ChrW(259) & ChrW(261), "LOWER_UNICODE_A_CHARS"  'lower case variants of the UNICODE character a
    d.Add ChrW(256) & ChrW(258) & ChrW(260), "UPPER_UNICODE_A_CHARS"  'upper case variants of the UNICODE character a
currentTest = "[DCM-2] Dictionary CompareMode - Retrieving items with binary compare mode"
    If d.item("keya") <> "itema" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.item("keyA") <> "itemA" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.item("keyB") <> "itemB" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.item("KeYB") <> "" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.item("KeYB") <> Empty Or d.item("KEYB") <> Empty Or d.item("keyb") <> Empty Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.item("KEYC") <> "ITEMC" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.item("KEyC") <> Empty Or d.item("Keyc") <> Empty Or d.item("keyc") <> Empty Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.item(Chr(128)) <> "EURO128" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.item(Chr(164)) <> "EURO164" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.item(ChrW(8352)) <> "EURO8352" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.item(Chr(128)) <> "EURO128" Then UtilAddArrayValue errors, ("-> " & currentTest)
    'If d.item("€") <> "EURO128" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.item(ChrW(183)) <> "BULLET1" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.item(ChrW(8226)) <> "BULLET2" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.item(ChrW(8729)) <> "BULLET3" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.item(ChrW(9679)) <> "BULLET4" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.item(ChrW(9702)) <> "BULLET5" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.item("aàáâãäå") <> "LOWER_ASCII_A_CHARS" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.item("AÀÁÂÃÄÅ") <> "UPPER_ASCII_A_CHARS" Then UtilAddArrayValue errors, ("-> " & currentTest) ' upper case keys will be sorted before lower case keys
    If d.item(ChrW(257) & ChrW(259) & ChrW(261)) <> "LOWER_UNICODE_A_CHARS" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.item(ChrW(256) & ChrW(258) & ChrW(260)) <> "UPPER_UNICODE_A_CHARS" Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[DCM-3] Dictionary CompareMode - Change of compare mode must not be possible if dictionary contains keys"
    On Error Resume Next: Err.Clear: d.CompareMode = 1 ' throws error, in DictCollection implementation this is possible
    If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest & ": Change of .CompareMode must throw error 5")
    Err.Clear
    If d.CompareMode <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.item("keya") <> "itema" Then UtilAddArrayValue errors, ("-> " & currentTest) ' <- In DictCollection implementation this would now be "itemA"
    If d.item("keyA") <> "itemA" Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear
currentTest = "[DCM-4] Dictionary CompareMode - Adding same keys lower/upper case with text compare must not be allowed"
    d.RemoveAll: d.CompareMode = 1
    On Error Resume Next
    Err.Clear: d.Add "keya", 1: d.Add "KEYA", 1: If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest & ": Adding ""keya"" twice must throw error 457")
    'Err.Clear: d.Add "aàáâãäå", 2: d.Add "AÀÁÂÃÄÅ", 2: If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest & ": Adding ""aàáâãäå"" twice must throw error 457")
    'Err.Clear: d.Add Chr(128), 3: d.Add "€", 3: If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest & ": Adding ""€"" twice must throw error 457")
    Err.Clear
currentTest = "[DCM-5] Dictionary CompareMode - Change of CompareMode with empty Dictionary must work"
    d.RemoveAll
    d.CompareMode = 0: If d.CompareMode <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest)
    d.CompareMode = 1: If d.CompareMode <> 1 Then UtilAddArrayValue errors, ("-> " & currentTest)
    d.CompareMode = 2: If d.CompareMode <> 2 Then UtilAddArrayValue errors, ("-> " & currentTest)
    d.CompareMode = 0 ' set back to default
    
currentTest = "[DKC-1] Dictionary Key Changing - Changing the key of existing items"
    d.RemoveAll: d.Add "a", "itemD": d.Add "b", "itemE": d.Add "c", "itemF"
    d.Key("a") = "d": d.Key("b") = "e": d.Key("c") = "f"
    ' STRANGE: d.Item("d") <> "itemD" returns true
    If (StrComp(d.Item("d"), "itemD") <> 0) Or (StrComp(d.Item("e"), "itemE") <> 0) Or (StrComp(d.Item("f"), "itemF") <> 0) Then UtilAddArrayValue errors, ("-> " & currentTest & ": string keys")
    d.Add -1, "itemD": d.Add 0, "itemE": d.Add 1, "itemF": d.Key(-1) = -100: d.Key(0) = 2: d.Key(1) = 100
    If (StrComp(d.Item(-100), "itemD") <> 0) Or (StrComp(d.Item(2), "itemE") <> 0) Or (StrComp(d.Item(100), "itemF") <> 0) Then UtilAddArrayValue errors, ("-> " & currentTest & ": number keys")
    ' changing object keys
    Set o1 = New DictCollection: Set o2 = New DictCollection: Set o3 = New DictCollection: Set o4 = New DictCollection: Set o5 = New DictCollection: Set o6 = New DictCollection
    o1.DefaultValueEnabled = False: o2.DefaultValueEnabled = False: o3.DefaultValueEnabled = False: o4.DefaultValueEnabled = False: o5.DefaultValueEnabled = False: o6.DefaultValueEnabled = False
    d.Add o1, "itemD": d.Add o2, "itemE": d.Add o3, "itemF"
    d.Key(o1) = o4: d.Key(o2) = o5: d.Key(o3) = o6
    If (StrComp(d.Item(o4), "itemD") <> 0) Or (StrComp(d.Item(o5), "itemE") <> 0) Or (StrComp(d.Item(o6), "itemF") <> 0) Then UtilAddArrayValue errors, ("-> " & currentTest & ": object keys")
currentTest = "[DKC-2] Dictionary Key Changing - Assigning keys already in use must throw error 457"
    Set o1 = New DictCollection: Set o2 = New DictCollection: o1.DefaultValueEnabled = False: o2.DefaultValueEnabled = False: o1.Add "o1", "o1": o2.Add "o2", "o2"
    d.RemoveAll: d.Add "a", "itemA": d.Add "b", "itemB": d.Add 1, "itemC": d.Add 2, "itemD": d.Add o1, "itemE": d.Add o2, "itemF"
    On Error Resume Next
    Err.Clear: d.Key("a") = "b": If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: d.Key("b") = "a": If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: d.Key(1) = 2: If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: d.Key(2) = "b": If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: d.Key(o1) = o2: If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: d.Key(o2) = "a": If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear
currentTest = "[DKC-3] Dictionary Key Changing - Changing number keys to same number but with different datatype should not have any effect"
    testdata = Array(Array(CInt(1), "Integer"), Array(CLng(1), "Long"), Array(CSng(1), "Single"), Array(CDbl(1), "Double"), Array(CCur(1), "Currency"), Array(CDec(1), "Decimal"), Array(#12/31/1899#, "Date"))
    For i = 0 To UBound(testdata)
        d.RemoveAll: d.Add testdata(i)(0), testdata(i)(1)
        For j = 0 To UBound(testdata)
            If j <> i Then
                d.Key(testdata(i)(0)) = testdata(j)(0): Keys = d.Keys
                If TypeName(Keys(0)) <> TypeName(testdata(i)(0)) Then
                    UtilAddArrayValue errors, ("-> " & currentTest & ": re-assignment with " & TypeName(testdata(j)(0)) & " changed original key datatype of " & TypeName(testdata(i)(0)))
                End If
            End If
        Next
    Next

currentTest = "[DNIA-1] Dictionary Nonexisting item access - Nonexisting keys must return empty variant"
    d.RemoveAll
    val1 = d.Item("nonexisting"): If Not IsEmpty(val1) Then UtilAddArrayValue errors, ("-> " & currentTest)
    val1 = d.Item(0): If Not IsEmpty(val1) Then UtilAddArrayValue errors, ("-> " & currentTest)
    
currentTest = "[DAE-1] Dictionary Access Errors - .Add without parameters must throw Error 450"
    On Error Resume Next
    Err.Clear: d.RemoveAll:  d.Add: If Err.Number <> 450 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear
    
currentTest = "[DAE-2] Dictionary Access Errors - Adding existing key must throw error 457"
    On Error Resume Next
    d.RemoveAll: Err.Clear: d.Add "1", 1: If Err.Number <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: String).")
    d.Add "1", 2: If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest & " (Keytype: String)")
    Err.Clear: d.Add CInt(1), 1: If Err.Number <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: Integer).")
    Err.Clear: d.Add CInt(1), 2: If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest & " (Keytype: Integer)")
    Err.Clear: d.Add CLng(2), 1: If Err.Number <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: Long).")
    Err.Clear: d.Add CLng(2), 2: If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest & " (Keytype: Long)")
    Err.Clear: d.Add CSng(3), 1: If Err.Number <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: Single).")
    Err.Clear: d.Add CSng(3), 2: If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest & " (Keytype: Single)")
    Err.Clear: d.Add CDbl(4), 1: If Err.Number <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: Double).")
    Err.Clear: d.Add CDbl(4), 2: If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest & " (Keytype: Double)")
    Err.Clear: d.Add CCur(5), 1: If Err.Number <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: Currency).")
    Err.Clear: d.Add CCur(5), 2: If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest & " (Keytype: Currency)")
    'Err.Clear: d.Add CDec(6), 1: If Err.Number <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: Decimal).")
    'Err.Clear: d.Add CDec(6), 2: If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest & " (Keytype: Decimal)")
    Err.Clear: d.Add CDate(7), 1: If Err.Number <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: Date).")
    Err.Clear: d.Add CDate(7), 2: If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest & " (Keytype: Date)")
    Err.Clear: d.Add CBool(1), 1: If Err.Number <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: Boolean).")
    Err.Clear: d.Add CBool(1), 2: If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest & " (Keytype: Boolean)")
    Err.Clear: d.Add Empty, 1: If Err.Number <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: Empty).")
    Err.Clear: d.Add Empty, 2: If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest & " (Keytype: Empty)")
    d.RemoveAll
    Err.Clear: d.Add "", 1: If Err.Number <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: String '').")
    Err.Clear: d.Add Empty, 2: If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest & " (Keytype: Empty with existing String '')")
    d.RemoveAll
    Err.Clear: d.Add Empty, 1: If Err.Number <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: Empty).")
    Err.Clear: d.Add "", 2: If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest & " (Keytype: String '' with existing Empty)")
    d.RemoveAll
    Err.Clear: d.Add "", 1: If Err.Number <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: String '').")
    Err.Clear: d.Add "", 2: If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest & " (Keytype: String '' with existing String '')")
    Set o1 = New DictCollection: Err.Clear: d.Add o1, 1: If Err.Number <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: Object).")
    d.Add o1, 2: If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest & " (Keytype: Object)")
    Err.Clear
    
currentTest = "[DAE-3] Dictionary Access Errors - Changing nonexisting key must throw Error 32811"
    On Error Resume Next
    d.RemoveAll: Err.Clear: d.Key("b") = "c": If Err.Number <> 32811 Then UtilAddArrayValue errors, ("-> " & currentTest & " (Keytype: String)")
    Err.Clear: d.Key(CInt(1)) = 2: If Err.Number <> 32811 Then UtilAddArrayValue errors, ("-> " & currentTest & " (Keytype: Integer)")
    Err.Clear: d.Add 1, "a": d.Key(2) = 2: If Err.Number <> 32811 Then UtilAddArrayValue errors, ("-> " & currentTest & " (Keytype: Integer)")
    Err.Clear: Set o1 = New DictCollection: o1.DefaultValueEnabled = False: d.Key(o1) = 2: If Err.Number <> 32811 Then UtilAddArrayValue errors, ("-> " & currentTest & " (Keytype: Object)")
    Err.Clear
    
    
ExitFunction:
    testScriptingDictionaryCompatibility = errors
End Function

Private Function testCollectionCompatibility(d)
    Dim errors, currentTest, i, j, r, Key, removeCount, indexToBeRemoved, indexToBeSwapped1, indexToBeSwapped2
    Dim test, Items, Keys, testdata, val1, val2
    Dim o1, o2, o3, o4, o5, o6, isDC
    errors = Array()
    
    ' set emulation
    Select Case TypeName(d)
        Case "DictCollection": d.EmulateCollection = True: isDC = True
        Case Else: isDC = False
    End Select

currentTest = "[CES-1] Collection Empty State - Count must be zero"
    If d.Count <> 0 Then UtilAddArrayValue errors, ("-> " & currentTest)
    
    For i = d.Count To 1 Step -1: d.Remove i: Next  ' remove all
    For test = 1 To 9
        Select Case test
            Case 1: ' String
                currentTest = "[CBF-1] Collection Basic Functions - Adding, accessing, changing and removing items with keys of datatype String"
                testdata = Array(Array("", "a"), Array("keyb", "b"), Array("keyC", "c"), Array("keyd", "d"), Array("KEYE", "e"), Array("kEYf", "f"))
            Case 2: ' Integer
                currentTest = "[CBF-2] Collection Basic Functions - Adding, accessing, changing and removing items of datatype Integer"
                testdata = Array(Array("a", CInt(-10000)), Array("b", CInt(-1)), Array("c", CInt(0)), Array("d", CInt(1)), Array("e", CInt(10000)), Array("f", CInt(32767)))
            Case 3: ' Long
                currentTest = "[CBF-3] Collection Basic Functions - Adding, accessing, changing and removing items of datatype Long"
                testdata = Array(Array("a", CLng(-100000000)), Array("b", CLng(-1)), Array("c", CInt(0)), Array("d", CLng(1)), Array("e", CLng(100000000)), Array("f", CLng(2147483647)))
            Case 4: ' Single
                currentTest = "[CBF-4] Collection Basic Functions - Adding, accessing, changing and removing items of datatype Single"
                testdata = Array(Array("a", CSng(-1000.001)), Array("b", CSng(-1.000001)), Array("c", CSng(0)), Array("d", CSng(1000.001)), Array("e", CSng(1.000001)), Array("f", CSng(3.402823E+38)))
            Case 5: ' Double
                currentTest = "[CBF-5] Collection Basic Functions - Adding, accessing, changing and removing items of datatype Double"
                testdata = Array(Array("a", CDbl(-100000.000000001)), Array("b", CDbl(-1.00000000000001)), Array("c", CDbl(0)), Array("d", CDbl(100000.000000001)), Array("e", CDbl(1.00000000000001)), Array("f", CDbl(4.94065645841247E-324)))
            Case 6: ' Currency
                currentTest = "[CBF-6] Collection Basic Functions - Adding, accessing, changing and removing items of datatype Currency"
                testdata = Array(Array("a", CCur(-100000000000.0001)), Array("b", CCur(-1.0001)), Array("c", CCur(0)), Array("d", CCur(1.0001)), Array("e", CCur(100000000000.0001)), Array("f", CCur(922337203685477.5)))
            Case 7: ' Date
                currentTest = "[CBF-8] Collection Basic Functions - Adding, accessing, changing and removing items of datatype Date"
                testdata = Array(Array("a", #1/1/100#), Array("b", #1/1/1899 10:00:01 AM#), Array("c", #1/1/1900#), Array("d", #12/31/2000 12:00:00 PM#), Array("e", #5/5/2100 11:00:00 PM#), Array("f", CDate(#12/31/9999 11:59:59 PM#)))
            Case 8: ' Object
                currentTest = "[CBF-9] Collection Basic Functions - Adding, accessing, changing and removing items of datatype Object"
                Set o1 = New DictCollection: Set o2 = New DictCollection: Set o3 = New DictCollection: Set o4 = New DictCollection: Set o5 = New DictCollection: Set o6 = New DictCollection
                o1.DefaultValueEnabled = False: o2.DefaultValueEnabled = False: o3.DefaultValueEnabled = False: o4.DefaultValueEnabled = False: o5.DefaultValueEnabled = False: o6.DefaultValueEnabled = False
                testdata = Array(Array("a", o1), Array("b", o2), Array("c", o3), Array("d", o4), Array("e", o5), Array("f", o6))
            Case 9: ' Mixed data types
                currentTest = "[CBF-11] Collection Basic Functions - Adding, accessing, changing and removing items of mixed datatypes"
                Set o1 = New DictCollection: Set o2 = New DictCollection: o1.DefaultValueEnabled = False: o2.DefaultValueEnabled = False
                testdata = Array(Array("", Empty), Array("object1", o1), Array("Number=1", CInt(1)), Array("String1", "1"), Array("Number=1.00001", 1.00001), Array("Number=-1.239847E-10", -1.239847E-10), Array("object2", o2), Array("Date='12/31/2012 11:12:00 PM'", #12/31/2012 11:12:00 PM#))
        End Select
        For i = 0 To UBound(testdata)
             If isDC Then d.Add2 testdata(i)(1), testdata(i)(0), , Null Else d.Add testdata(i)(1), testdata(i)(0)
        Next
        removeCount = UBound(testdata)
        For r = 0 To removeCount
            If d.Count <> UBound(testdata) + 1 Then UtilAddArrayValue errors, ("-> " & currentTest & ": .Count must equal " & UBound(testdata) + 1)
            For i = 0 To UBound(testdata)
                ' compare type of .Item using index
                val1 = TypeName(d.Item(i + 1)): val2 = TypeName(testdata(i)(1))
                If val1 <> val2 Then
                    UtilAddArrayValue errors, ("-> " & currentTest & ": Access with index failed - .Item(" & i + 1 & ") is " & val1 & " but must be of type " & val2)
                ElseIf IsObject(d.Item(i + 1)) Then
                    If Not (d.Item(i + 1) Is testdata(i)(1)) Then UtilAddArrayValue errors, ("-> " & currentTest & ": Access with index failed - .Item(" & i + 1 & ") must be same object as testdata index=" & i)
                ElseIf d.Item(i + 1) <> testdata(i)(1) Then
                    UtilAddArrayValue errors, ("-> " & currentTest & ": Access with index failed - .Item(" & i + 1 & ") must be " & testdata(i)(1))
                End If
                
                If TypeName(testdata(i)(0)) = "String" Then
                    ' compare type of .Item using key
                    val1 = TypeName(d.Item(testdata(i)(0))): val2 = TypeName(testdata(i)(1))
                    If val1 <> val2 Then
                        UtilAddArrayValue errors, ("-> " & currentTest & ": Access with key failed - .Item(" & testdata(i)(0) & ") is " & val1 & " but must be of type " & val2)
                    ElseIf IsObject(d.Item(testdata(i)(0))) Then
                        If Not (d.Item(testdata(i)(0)) Is testdata(i)(1)) Then UtilAddArrayValue errors, ("-> " & currentTest & ": Access with key failed - .Item(" & i + 1 & ") must be same object as testdata index=" & i)
                    ElseIf d.Item(testdata(i)(0)) <> testdata(i)(1) Then
                        UtilAddArrayValue errors, ("-> " & currentTest & ": Access with key failed - .Item(" & i + 1 & ") must be " & testdata(i)(1))
                    End If
                    ' compare access with uppercase key
                    Key = UCase(testdata(i)(0))
                    If IsObject(d.Item(Key)) Then
                        If Not (d.Item(Key) Is testdata(i)(1)) Then UtilAddArrayValue errors, ("-> " & currentTest & ": Access with uppercase key failed - .Item(" & Key & ") must be " & testdata(i)(1))
                    ElseIf d.Item(Key) <> testdata(i)(1) Then
                        UtilAddArrayValue errors, ("-> " & currentTest & ": Access with uppercase key failed - .Item(" & Key & ") must be " & testdata(i)(1))
                    End If
                    ' compare access with lowercase key
                    Key = LCase(testdata(i)(0))
                    If IsObject(d.Item(Key)) Then
                        If Not (d.Item(Key) Is testdata(i)(1)) Then UtilAddArrayValue errors, ("-> " & currentTest & ": Access with lowercase key failed - .Item(" & Key & ") must be " & testdata(i)(1))
                    ElseIf d.Item(Key) <> testdata(i)(1) Then
                        UtilAddArrayValue errors, ("-> " & currentTest & ": Access with lowercase key failed - .Item(" & Key & ") must be " & testdata(i)(1))
                    End If
                Else
                    ' keys with type other than "String" must throw errors
                    Err.Clear: On Error Resume Next
                    val1 = d.Item(testdata(i)(0)): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest & ": Access with non-string key did not throw error!")
                    Err.Clear
                End If
            Next
            ' identify index to be removed, and indexes to be swapped
            indexToBeSwapped1 = -1: indexToBeSwapped1 = -1 ' do not swap anything
            Select Case r
                Case 0, 1:
                    indexToBeRemoved = UBound(testdata) \ 2 ' take item from the middle
                    If UBound(testdata) >= 3 Then indexToBeSwapped1 = indexToBeRemoved + 1: indexToBeSwapped2 = indexToBeRemoved - 1 ' swap values above and below index
                Case Else:
                    If UBound(testdata) Mod 2 = 0 Then
                        indexToBeRemoved = 0 ' take first index
                        If UBound(testdata) >= 2 Then indexToBeSwapped1 = 1: indexToBeSwapped2 = 2 ' swap others above
                    Else
                        indexToBeRemoved = UBound(testdata) ' take last
                        If UBound(testdata) >= 2 Then indexToBeSwapped1 = indexToBeRemoved - 1: indexToBeSwapped2 = indexToBeRemoved - 2 ' swap others below
                    End If
            End Select
            If indexToBeSwapped1 > -1 And indexToBeSwapped2 > -1 Then
                ' swap values in test array
                If IsObject(testdata(indexToBeSwapped1)(1)) Then Set testdata(indexToBeRemoved)(1) = testdata(indexToBeSwapped1)(1) Else testdata(indexToBeRemoved)(1) = testdata(indexToBeSwapped1)(1)
                If IsObject(testdata(indexToBeSwapped2)(1)) Then Set testdata(indexToBeSwapped1)(1) = testdata(indexToBeSwapped2)(1) Else testdata(indexToBeSwapped1)(1) = testdata(indexToBeSwapped2)(1)
                If IsObject(testdata(indexToBeRemoved)(1)) Then Set testdata(indexToBeSwapped2)(1) = testdata(indexToBeRemoved)(1) Else testdata(indexToBeSwapped2)(1) = testdata(indexToBeRemoved)(1)
                ' remove and re-add items at new indexes
                If TypeName(testdata(indexToBeSwapped1)(0)) = "String" Then d.Remove testdata(indexToBeSwapped1)(0) Else d.Remove indexToBeSwapped1 + 1
                If TypeName(testdata(indexToBeSwapped2)(0)) = "String" Then d.Remove testdata(indexToBeSwapped2)(0) Else d.Remove indexToBeSwapped2 + 1
                ' ensure index1 is smaller than index2
                If indexToBeSwapped1 > indexToBeSwapped2 Then val1 = indexToBeSwapped2: indexToBeSwapped2 = indexToBeSwapped1: indexToBeSwapped1 = val1
                If isDC Then
                    If indexToBeSwapped1 = 0 Then
                        d.Add2 testdata(indexToBeSwapped1)(1), testdata(indexToBeSwapped1)(0), indexToBeSwapped1 + 1, Null ' add before
                    Else
                        d.Add2 testdata(indexToBeSwapped1)(1), testdata(indexToBeSwapped1)(0),, indexToBeSwapped1 ' add after
                    End If
                    If indexToBeSwapped2 = 0 Then
                        d.Add2 testdata(indexToBeSwapped2)(1), testdata(indexToBeSwapped2)(0), indexToBeSwapped2 + 1, Null ' add before
                    Else
                        d.Add2 testdata(indexToBeSwapped2)(1), testdata(indexToBeSwapped2)(0),, indexToBeSwapped2 ' add after
                    End If
                Else
                    If indexToBeSwapped1 = 0 Then
                        d.Add testdata(indexToBeSwapped1)(1), testdata(indexToBeSwapped1)(0), indexToBeSwapped1 + 1, Null ' add before
                    Else
                        d.Add testdata(indexToBeSwapped1)(1), testdata(indexToBeSwapped1)(0),, indexToBeSwapped1 ' add after
                    End If
                    If indexToBeSwapped2 = 0 Then
                        d.Add testdata(indexToBeSwapped2)(1), testdata(indexToBeSwapped2)(0), indexToBeSwapped2 + 1, Null ' add before
                    Else
                        d.Add testdata(indexToBeSwapped2)(1), testdata(indexToBeSwapped2)(0),, indexToBeSwapped2 ' add after
                    End If
                End If
            End If
            If TypeName(testdata(indexToBeRemoved)(0)) = "String" Then d.Remove testdata(indexToBeRemoved)(0) Else d.Remove indexToBeRemoved + 1
            UtilRemoveArrayValueByIndex testdata, indexToBeRemoved ' remove 1 item in test array
        Next
    Next
    
currentTest = "[CIA-1] Collection Index Access"
    For i = d.Count To 1 Step -1: d.Remove i: Next  ' remove all
    If isDC Then d.Add2 "a",,, Null: d.Add2 "b",,, Null Else d.Add "a": d.Add "b"
    If d.Item(CInt(1)) <> "a" Or d.Item(CInt(2)) <> "b" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.Item(CLng(1)) <> "a" Or d.Item(CLng(2)) <> "b" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.Item(CSng(1.4)) <> "a" Or d.Item(CSng(1.6)) <> "b" Or d.Item(CSng(1.5)) <> "b" Or d.Item(CSng(2)) <> "b" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.Item(CDbl(1.4)) <> "a" Or d.Item(CDbl(1.6)) <> "b" Or d.Item(CDbl(1.5)) <> "b" Or d.Item(CDbl(2)) <> "b" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.Item(CCur(1.4)) <> "a" Or d.Item(CCur(1.6)) <> "b" Or d.Item(CCur(1.5)) <> "b" Or d.Item(CCur(2)) <> "b" Then UtilAddArrayValue errors, ("-> " & currentTest)
    'If d.Item(CDec(1.4)) <> "a" Or d.Item(CDec(1.6)) <> "b" Or d.Item(CDec(1.5)) <> "b" Or d.Item(CDec(2)) <> "b" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If d.Item(CDate(1.4)) <> "a" Or d.Item(CDate(1.6)) <> "b" Or d.Item(CDate(1.5)) <> "b" Or d.Item(CDate(2)) <> "b" Then UtilAddArrayValue errors, ("-> " & currentTest)
        

currentTest = "[CCIA-1] Collection Case Insensitive Access"
    For i = d.Count To 1 Step -1: d.Remove i: Next  ' remove all
    If isDC Then
        d.Add2 "LOWER_A", "a", , Null 'lower case variants of the ASCII character a
    Else
        d.Add "LOWER_A", "a" 'lower case variants of the ASCII character a
    End If
    If d.Item("A") <> "LOWER_A" Then UtilAddArrayValue errors, ("-> " & currentTest)
    If isDC Then
        d.Add2 "UNICODE_A_CHARS", ChrW(257) & ChrW(259) & ChrW(261), , Null 'lower case variants of the ASCII character a
    Else
        d.Add "UNICODE_A_CHARS", ChrW(257) & ChrW(259) & ChrW(261) 'lower case variants of the ASCII character a
    End If
    If d.Item(ChrW(256) & ChrW(258) & ChrW(260)) <> "UNICODE_A_CHARS" Then UtilAddArrayValue errors, ("-> " & currentTest)


currentTest = "[CAE-1] Collection Access Errors - Accessing nonexisting keys in empty Collection must throw Error 5 'Invalid procedure call or argument'"
    For i = d.Count To 1 Step -1: d.Remove i: Next  ' remove all
    On Error Resume Next
    Err.Clear: val1 = d.Item(Empty): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item(""): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item("a"): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[CAE-2] Collection Access Errors - Accessing nonexisting indexes in empty Collection must throw Error 5 'Invalid procedure call or argument'"
    Err.Clear: val1 = d.Item(0): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item(1): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item(-1): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item(1000): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[CAE-3] Collection Access Errors - Removing nonexisting indexes in empty Collection must throw Error 5 'Invalid procedure call or argument'"
    Err.Clear: val1 = d.Remove(0): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Remove(1): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Remove(-1): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Remove(1000): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[CAE-4] Collection Access Errors - Removing nonexisting keys must in empty Collection throw Error 5 'Invalid procedure call or argument'"
    Err.Clear: val1 = d.Remove(Empty): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Remove(""): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Remove("a"): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[CAE-5] Collection Access Errors - Accessing invalid keytypes in empty Collection must throw Error 5 'Invalid procedure call or argument'"
    Err.Clear: val1 = d.Remove(GetMissingValue(,1)): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item(Null): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item(Nothing): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
    'Err.Clear: val1 = d.Item(CVErr(11)): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[CAE-6] Collection Access Errors - Accessing nonexisting keys in filled Collection must throw Error 9 'Subscript out of range'"
    For i = d.Count To 1 Step -1: d.Remove i: Next  ' remove all
    If isDC Then d.Add2 "a",,, Null Else d.Add "a"
    Err.Clear: val1 = d.Item(Empty): If Err.Number <> 9 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[CAE-7] Collection Access Errors - Accessing nonexisting String keys in filled Collection must throw Error 5 'Invalid procedure call or argument'"
    Err.Clear: val1 = d.Item(""): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item("b"): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[CAE-8] Collection Access Errors - Accessing nonexisting indexes in filled Collection must throw Error 9 'Subscript out of range'"
    Err.Clear: val1 = d.Item(0): If Err.Number <> 9 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item(2): If Err.Number <> 9 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item(-1): If Err.Number <> 9 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item(1000): If Err.Number <> 9 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[CAE-9] Collection Access Errors - Removing nonexisting indexes in filled Collection must throw Error 9 'Subscript out of range'"
    Err.Clear: val1 = d.Remove(0): If Err.Number <> 9 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Remove(3): If Err.Number <> 9 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Remove(-1): If Err.Number <> 9 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Remove(1000): If Err.Number <> 9 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[CAE-10] Collection Access Errors - Removing nonexisting keys in filled Collection must throw Error 5 'Invalid procedure call or argument'"
    Err.Clear: val1 = d.Remove(""): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Remove("a"): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[CAE-11] Collection Access Errors - Removing key Empty in filled Collection must throw Error 9 'Subscript out of range'"
    Err.Clear: val1 = d.Remove(Empty): If Err.Number <> 9 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[CAE-12] Collection Access Errors - Removing key of type Missing in filled Collection must throw Error 5 'Invalid procedure call or argument'"
    Err.Clear: val1 = d.Remove(GetMissingValue(,1)): If Err.Number <> 5 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[CAE-13] Collection Access Errors - Accessing keys of invalid datatype in filled Collection must throw Error 13 'Type mismatch'"
    Err.Clear: val1 = d.Item(Null): If Err.Number <> 13 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item(Nothing): If Err.Number <> 13 Then UtilAddArrayValue errors, ("-> " & currentTest)
    'Err.Clear: val1 = d.Item(CVErr(11)): If Err.Number <> 13 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[CAE-14] Collection Access Errors - Adding existing keys must throw Error 457 'This key is already associated with an element of this collection'"
    For i = d.Count To 1 Step -1: d.Remove i: Next  ' remove all
    If isDC Then d.Add2 "itemA", "keyA",, Null: d.Add2 "itemB", "",, Null Else d.Add "itemA", "keyA": d.Add "itemB", ""
    Err.Clear: If isDC Then d.Add2 "itemC", "keya",, Null Else d.Add "itemC", "keya"
        If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: If isDC Then d.Add2 "itemC", "",, Null Else d.Add "itemC", ""
        If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[CAE-15] Collection Access Errors - Adding keys of invalid datatype must throw Error 13 'Type mismatch'"
    For i = d.Count To 1 Step -1: d.Remove i: Next  ' remove all
    testdata = Array(CInt(1), CLng(1), CSng(1), CDbl(1), CCur(1), CDec(1), CDate(1), Null, Nothing)
    For i = 0 To UBound(testdata)
        Err.Clear
        If isDC Then d.Add2 "item", testdata(i),, Null Else d.Add "item", testdata(i)
        If Err.Number <> 13 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Next
currentTest = "[CAE-16] Collection Access Errors - Inserting items with before or after invalid index must throw Error 9 'Subscript out of range'"
    For i = d.Count To 1 Step -1: d.Remove i: Next  ' remove all
    If isDC Then d.Add2 "itemA",,, Null: d.Add2 "itemB",,, Null Else d.Add "itemA": d.Add "itemB" ' index 1,2
    Err.Clear: If isDC Then d.Add2 "itemC",, 0, Null Else d.Add "itemC",, 0 ' before -1
        If Err.Number <> 9 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: If isDC Then d.Add2 "itemC",, 0, Null Else d.Add "itemC",, 0 ' before 0
        If Err.Number <> 9 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: If isDC Then d.Add2 "itemC",, 3, Null Else d.Add "itemC",, 3 ' before 3
        If Err.Number <> 9 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: If isDC Then d.Add2 "itemC",,, 0 Else d.Add "itemC",, 0 ' after 0
        If Err.Number <> 9 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: If isDC Then d.Add2 "itemC",,, 3 Else d.Add "itemC",, 3 ' after 3
        If Err.Number <> 9 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[CAE-17] Collection Access Errors - Adding items with existing indexes as keys must throw Error 457 'This key is already associated with an element of this collection'"
    For i = d.Count To 1 Step -1: d.Remove i: Next  ' remove all
    If isDC Then d.Add2 "itemA",,,Null: d.Add2 "itemB",,,Null Else d.Add "itemA": d.Add "itemB" ' index 1,2
    Err.Clear: If isDC Then d.Add2 "itemC", 1,, Null Else d.Add "itemC", 1 ' index=1
        If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest)
    Err.Clear: If isDC Then d.Add2 "itemC", 2,, Null Else d.Add "itemC", 2 ' index=2
        If Err.Number <> 457 Then UtilAddArrayValue errors, ("-> " & currentTest)
currentTest = "[CAE-18] Collection Access Errors - Adding items using invalid datatye for Before or After must throw Error 5, 9 or 13"
    For i = d.Count To 1 Step -1: d.Remove i: Next  ' remove all
    Set o1 = New DictCollection: o1.DefaultValueEnabled = False
    testdata = Array(Array("a", 5), Array(Empty, 9), Array(CVErr(1), 13), Array(o1, 13), Array(Nothing, 13), Array(Null, 13))
    If isDC Then d.Add2 "itemA": d.Add2 "itemB",,, Null Else d.Add "itemA": d.Add "itemB" ' index 1,2
    For i = 0 To UBound(testdata)
        Err.Clear: If isDC Then d.Add2 "itemC",, testdata(i)(0), Null Else d.Add "itemC",, testdata(i)(0) ' before = invalid datatype
        If Err.Number <> testdata(i)(1) Then UtilAddArrayValue errors, ("-> " & currentTest)
        Err.Clear: If isDC Then d.Add2 "itemC",,, testdata(i)(0) Else d.Add "itemC",,, testdata(i)(0) ' after = invalid datatype
        If Err.Number <> testdata(i)(1) Then UtilAddArrayValue errors, ("-> " & currentTest)
    Next    
    Err.Clear

' Adding items with Missing as key is OK but accessing them with Missing throws Error 5 'Invalid procedure call or argument'

    testCollectionCompatibility = errors
End Function
' ======================== END: COMPATIBILITY TESTS ===============================================================================================
End Class

' ======================== START: VBA INFO ========================================================================================================
'Data Type          Bytes   Declaration             Value Range
'--------------------------------------------------------------
'Boolean            2                               True or False
'Integer            2       100% 1e2%               -32,768 to 32767
'Long               4       -65000& -6.5e4&         -2,147,483,648 to 2,147,483,647
'Single             4       -0.17! -1.7E-1!         -3.402823E38 to 3.402823E38
'Double (negative)  8       -0.0003# -3E-4#         -1.79769313486232E308 to -4.94065645841247E-324
'Double (positive)  8       1.00000000000001#       4.94065645841247E-324 to 1.79769313486232E308
'Currency           8       10.237465345@           -922,337,203,685,477.5808 to 922,337,203,685,477.5807
'Date (Double)      8       #7/7/2009 11:00:00 PM#  "1/1/100" to "31/12/9999 11:59:59 PM"
'String             1       "a" "ab" ""             per character Varies according to the number of characters
'Object             4                               Any defined object
'Variant            varies                          Any data type
'Used defined       varies                          varies
' ======================== END: VBA INFO ==========================================================================================================

' ======================== START: VBSCRIPT INFO ===================================================================================================
' - Dim cannot use types, everything is Variant/Integer = -1 by default
' - Goto is not supported, 'On Error Goto 0' for restoring break-on-error mode does not exist
' - optional parameters and default values for parameters are not supported
' - the last parameter of function call cannot be Missing (except with a workaround function that returns a missing value)
' - Expressing numbers as currency type with 100.001@ is not supported
' - Expressing numbers as double type with 100.001# is not supported
' - Currency Datatype maximum value is 922337203685477.5 instead of 922337203685477.5807 in VBA
' - Decimal data type is not supported, CDec() throws error
' - Arrays have to be declared like 'Dim MyArr()' or initialized using a = Array() before typical array operations ban be performed
' - CVErr() cannot be used as in VBA, throws error
' - IsMissing() does not exist, has to be implemented as (VarType(p) = vbError)
' - iif(cond,truepart,falsepart) does not exist, has to implemented as if ... then ... else
' - implicit casts have to be made explicit using CInt(), CLng(), etc.
' - Class default property is declared as 'Public Default Property Get MyProperty()'
' ======================== END: VBSCRIPT INFO =====================================================================================================
