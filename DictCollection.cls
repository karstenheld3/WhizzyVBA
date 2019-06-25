VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DictCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ======================== START: DOCUMENTATION =====================================================================================================
' DictCollection 041 (VBA Version) - A mix between Dictionary and Collection that can emulate both
' Karsten Held 2016-2019, karstenheld3@gmail.com, MIT License
'
' ----------- HOW TO USE ----------------------------------------------------------------------------------------------------------------------------
' Create an empty collection:                       Set dc = New DictCollection
' Add an item with a key:                           dc.Add "key1", 123.45                           -> dc.Item(0) = 123.45
' Add an item without a key:                        dc.Add , "A"                                    -> dc.Item(1) = "A"
' Get the number or items:                          n = dc.Count                                    -> n = 2
' Get the number or keys:                           k = dc.KeyCount                                 -> k = 1
' Get an item by key:                               a = dc.Item("key1")                             -> a = 123.45
' Get an item by index:                             a = dc.Item(0)                                  -> a = 123.45
'                                                   b = dc.Item(1)                                  -> b = "A"
' dc() is shorthand for .Item():                    b = dc(1)                                       -> b = "A"
' Set an item by index:                             dc(0) = 100                                     -> dc.Item(0) = 100
' Set an item by key:                               dc("key1") = 200                                -> dc.Item(0) = 200
' Add an item by key:                               dc("key2") = "B"                                -> dc.Item(2) = "B"
' Nesting another DictCollection:                   dc.Add "key3", new DictCollection               -> dc.Item(3) = [Empty DictCollection]
' Add to nested DictCollection:                     dc("key3").Add "key4", 456.78                   -> dc.Item(3).Item(0) = 456.78
'                                                   dc("key3").Add , "C"                            -> dc.Item(3).Item(1) = "C"
' Get items from nested DictCollection:             c = dc("key3")("key4")                          -> c = 456.78
'                                                   d = dc("key3")(1)                               -> d = "C"
' Check if nested item exists:                      dc("key3")(1) = dc.NonExistingValue             -> false
'   ('endless subcollection chaining')              dc("key3")("wrongkey") = dc.NonExistingValue    -> true
' Remove an item by key:                            dc.Remove("key1")                               -> "A" is now at index 0
' Remove an item by index:                          dc.Remove(0)                                    -> "B" is now at index 0
' Get the index of an item by key:                  idx = dc.IndexOfKey("key3")                     -> idx = 1
' Get an item's key (Empty if none):                a = dc.KeyOfItemAt(0)                           -> a = "key2"
' Check if key exists:                              check = dc.Exists("key3")                       -> check = True
' Change an items key:                              dc.Key("key2") = "key4"                         -> "key2" changed to "key4"
' Change an items key while removing another
'   item that used this key before:                 dc.Key("key4") = "key3"                         -> "key4" -> "key3", dc.Item(1) removed
' Get keys for each item incl. Empty:               ka = dc.Keys                                    -> ka = Array("key3")
' Get all items in the added order:                 ia = dc.Items                                   -> ia = Array("B")
' Insert an item with key at index=0:               dc.Insert "D", 0, "key5"                        -> dc.Item(0) = "D", dc.Item(1) = "B"
' Insert an item without key at index=3:            dc.Insert "E", 3, ""                            -> dc.Item(2) = Empty, dc.Item(3) = "E"
' Get all items with empty values:                  ia = dc.Items                                   -> ia = Array("D", "B", Empty, "E")
' Get all items without empty values:               ia = dc.Items(False)                            -> ia = Array("D", "B", "E")
' Get all used keys in item order:                  ka = dc.Keys(False)                             -> ka(0) = Array("key5","key3")
' Get all keys in sorted order:                     ka = dc.SortedKeys                              -> ka = Array("key3", "key5")
' Remove all items from DictCollection:             dc.RemoveAll
' Move an item to another index:                    dc.Move(fromIndex, toIndex)
' Get an item by index (fast)                       a = dc.ItemAt(Index)
' Set an item by index (fast)                       dc.ItemAt(Index) = a
' Get an item by key (fast)                         a = dc.ItemOf("key1")
' Set an item by key (fast)                         dc.ItemOf("key1") = a
' Add (insert) a DictCollection                     dc.AddDC("a",0) -> at index = 0 or dc.AddDC("a") -> after last item
' Add multiple SubCollections to parent             dc.AddDC("b").AddDC("c").AddDC("d")
' Get nested DictCollections                        dc.AsDC("a").AsDC("a-a") -> will return nested DictCollection or "[NONEXISTING]"
' Get or Add nested DictCollections                 dc.SubDC("subkey1").SubDC("subsubkey1") -> will create nonexisting subkeys as DictCollections
'
' ----------- GLOBAL SETTINGS ------------------------------------------------------------------------------------------------------------------------
' dc.ThowErrors = False (default)                   will return .NonExistingValue when accessing nonexistent indexes or keys
' dc.ThowErrors = True                              will throw errors when accessing nonexistent indexes or keys
' dc.LazySorting = True (default)                   will sort key array at first access using a key and thus speed up .Add
' dc.LazySorting = False                            will sort key array at every .Add if key is used
' dc.CompareMode = 0 (default)                      keys are case-sensitive, will binary compare keys, fastest compare method
' dc.CompareMode = 1                                keys are case-insensitive, ƒ = ‰ = A = a = · = ¡
' dc.CompareMode = 2                                keys are case-insensitive (only MSAccess), uses localeID for matching
' dc.EmulateDictionary = True                       DictCollection will behave like Scripting.Dictionary (no indexes, just keys)
' dc.EmulateCollection = True                       DictCollection will behave like VB Collection (Collection.Add implemented as DictCollection.Add2)
' dc.ZeroBasedIndex = True (default)                first item can be accessed with index = 0 (when ZeroBasedIndex=False, first index = 1)
' dc.DefaultValueEnabled = True (default)           calling default property without argument returns either the first item, "[EMPTY]", or "[NONEXISTING]"
'   -> see test [DVS-1]                             allows endless subcollection chaining: 'set dc2 = dc1("a")("nonexisting")("b")' (ThrowErrors=False)
' dc.DefaultValueEnabled = False                    calling default property without argument as dc or dc() returns a reference to the DictCollection
'   -> see test [DVS-2]                             allows it to be used like a regular object: 'VarType(dc)' and 'dc Is DictCollection' will work
' dc.NonExistingValue = "[NONEXISTING]" (default)   the value to be returned for nonexisting items if ThrowError=False
'                                                   can be used to overwrite this value with anything except objects, e.g. Empty, "", 0, False
' dc.EmptyCollectionValue = "[EMPTY]" (default)     the default value of empty DictCollections (ThrowError=False, DefaultValueEnabled=True)
'                                                   can be used to overwrite this value with anything except objects, e.g. Empty, "", 0, False
' dc.CollectionType sets/gets the collection type:  0 = NonExisting, 1 = Empty Array, 2 = Filled Array, 3 = Empty Key-Value-Store,
'                                                   4 = Filled Key-Value-Store, 5 = Key-Value-Store with at least one item having no key
'
' ----------- OTHER DICTCOLLECTION FUNCTIONS ----------------------------------------------------------------------------------------------------------
'  * = function is chainable, it always returns the DictCollection which it is a members of
' Collection-compatible Add function                * dc.Add2(Item, [Key], [Before], [After]) As DictCollection
' Add key-value pairs from 2-dimensional            * dc.AddPairs(Arr, [KeyIsFirstColumn]) As DictCollection
'   or nested array (doesn't add empty keys)          Example: dc.AddPairs Array(Array("a",1), Array("b",2))
' Copy all items and keys to another DC             * dc.CopyItems(TargetCollection, [TargetIndex]) As DictCollection
' Copy all items and keys to new DictCollection     dc.CopyItemsToNew() As DictCollection
' Copy settings from on DC to another DC            dc.CopyAllSettingsExceptCollectionType(FromCollection, ToCollection)
' Copy subcollection chaining settings              dc.CopySubCollectionChainingSettings(FromCollection, ToCollection)
' Clone entire DictCollection tree to new DC        dc.CloneToNew() as DictCollection
' Clone entire DictCollection to existing DC        * dc.CloneTo(TargetCollection) as DictCollection
' Clone to existing DC and keep its content         * dc.CloneToPreserve(TargetCollection) as DictCollection
' Create new DC by dropping specified keys          dc.DropKeysToNew(KeyArr, [AllowWildcards], [CompareMode]) As DictCollection
' Create new DC by keeping specified keys           dc.KeepKeysToNew(KeyArr, [AllowWildcards], [CompareMode]) As DictCollection
' Create new DC by matching specified keys          dc.MatchKeysToNew(KeyArr, [NotFoundvalue], [CompareMode]) As DictCollection
' Set item by index or key (chainable function)     * dc.SetItem(IndexOrKey, Value) As DictCollection
' Increment stored number by 1 or Amount            dc.Increment(IndexOrKey, [Amount]) As Variant
' Check if item has key                             dc.ItemHasKey(ItemIndex) As Boolean
' Check if item is a DictCollection                 UNTESTED dc.ItemIsDC(IndexOrKey) As Boolean
' Check if item is Object                           UNTESTED dc.ItemIsObject(IndexOrKey) As Boolean
' Get item index associated with key (String)       UNTESTED dc.IndexOfKey2(Key) As Long
' Get internal key storage index of an items key    dc.KeyIndexOfItemAt(Index) As Long
' Get key at internal key storage index             dc.KeyAtKeyIndex(KeyIndex) As Variant
' Get string key at internal key storage index      dc.KeyAtKeyIndexAsString(KeyIndex) As Variant
' Get an items string key from internal key storage dc.KeyOfItemAtAsString(Index) As String
' Get all items and keys as Array(i,k)              UNTESTED dc.ToArray() As Variant
' Get items and keys as Array(i,k)                  UNTESTED dc.ItemsAndKeys([IncludeEmptyItems], [IncludeItemsWithoutKeys]) As Variant
' Get items and keys as Array(i,k) sorted by keys   UNTESTED dc.ItemsAndKeysSortedByKeys([IncludeEmptyItems], [IncludeItemsWithoutKeys]) As Variant
' Get items as Array sorted by keys                 UNTESTED dc.ItemsSortedByKeys([IncludeEmptyItems], [IncludeItemsWithoutKeys]) As Variant
' Get keys and items as Array(k,i) sorted by keys   UNTESTED dc.SortedKeysAndItems([IncludeEmptyKeys], [IncludeEmptyItems]) As Variant
' Get subCol values by key/column as Array(r,c)     UNTESTED dc.ToColumns(ColHeaderKeys, [FlattenSubCollections]) As Variant
' Get subCols as Array(r,c) incl. col/row keys      UNTESTED dc.ToTable(FillValue, [FlattenSubCollections]) As Variant
' Find key index where insert leaves keys sorted    UNTESTED dc.FindKeyInsertIndex(SearchedKey As Variant, [CompareMode]) As Long
' Find all keys that start with text (as Array)     UNTESTED dc.FindKeysThatStartWith(SearchText, [CompareMode]) As Variant
' Get infos about DictCollection tree content       dc.AnalyzeDictCollectionTree([Recursive]) As DictCollection
' Assign "_POS" keys to items that have no keys     dc.EnsureAllItemsHaveKeys()
' Flatten DC tree to single key-value-store         * dc.Flatten() As DictCollection
' Flatten DC tree to new DictCollection             dc.FlattenToNew() As DictCollection
' Restore DC tree from single key-value-store       * dc.Unflatten() As DictCollection
' Restore DC tree to new DictCollection             dc.UnflattenToNew() As DictCollection
' Demo basic functionality in Immediate Window      dc.DemoBasicFunctionality()
' Run all tests with output to Immediate Window     dc.SelfTest([DebugPrint])
' Run compatibility tests and return errors         dc.TestCompatibility([DebugPrint]) as Variant
' Run functionality tests and return errors         dc.TestFunctionality([DebugPrint]) as Variant
'
' ----------- OTHER UTILITY FUNCTIONS ----------------------------------------------------------------------------------------------------------------
' Get missing argument (Error 448) as Variant       dc.UtilGetMissingValue([DoNotPassAnythingHere])
' Assigns object or value to Variant variable       dc.UtilAssignFromTo(FromVariable, ToVariable)
' Add value to array (creates one)                  UNTESTED dc.UtilAddArrayValue(Arr As Variant, Val As Variant)
' Find index of value in array or return -1         dc.UtilFindArrayIndex(Arr, Val) As Long
' Remove a value from an array by index             dc.UtilRemoveArrayValueByIndex(Arr, Index)
' Get array dimensions (0 = uninitialized)          dc.UtilArrayDimensions(Arr) As Integer
' Remove a value from an array                      dc.UtilRemoveArrayValue(Arr, Val)
' Converts array to a one-dimensional array         UNTESTED dc.UtilArrayConvertTo1Dim(Arr) As Variant
' Converts array to two-dimensional array           UNTESTED dc.UtilArrayConvertTo2Dim(Arr, RowCount, ColumnCount, FillValue) As Variant
' Sort one/two-dimensional/nested array             UNTESTED dc.UtilSortArray(Arr, FromIndex, ToIndex)
' Sort one/two-dimensional/nested array (StrComp)   UNTESTED dc.UtilSortStringArray(Arr, FromIndex, ToIndex, CompareMode)
' Check if text has only number chars               UNTESTED dc.UtilStringConsistsOfNumericAsciiChars(Text) As Boolean
' Build concatenated string by repeating a text     UNTESTED dc.UtilStringRepeat(Text, NumberOfTimes) As String
' Check if text starts with another text            UNTESTED dc.UtilStringStartsWith(Text, SearchText, [CompareMode]) As Boolean
' Find first occurrence of chars within text        UNTESTED dc.UtilStringFindChars(InputString, StartPos, Chars As String) As Long
' Check if text matches splitted wildcard ("*.*")   dc.UtilStringMatchesWildcard(TextToCheck, WildcardSplit, [CompareMode]) As Boolean
'
' ----------- INFORMATION ---------------------------------------------------------------------------------------------------------------------------
' - stored items can be Empty, Nothing, Null, Error, Array, Object, Integer, Long, Single, Double, Currency, Decimal, Date, Byte
' - stored keys can be only nonempty Strings; .EmulateDictionary=True allows keys of all data types; .EmulateCollection=True allows "" as key
' - auto-expansion (.ThrowErrors=false): assigning an item to an index greater than the item storage size will expand the storage to that index;
'   implicitly created items will be Empty; Example: Set dc = new DictCollection: dc(2) = "a" -> dc.Items = Array(Empty,Empty,"a")
' - with .EmulateDictionary=True .DefaultProperty(), .Item(), .IndexOfKey(), .KeyOfItemAt() support keys of all data types
' - .ItemOf(), .IndexOfKey2() support only keys of type String and are slightly faster (they also work when emulating Dictionary or Collection)
' - with .EmulateDictionary=True the property .Item(IndexOrKey) will interpret any IndexOrKey argument as a key; use .ItemAt(Index) for index access
' - with .EmulateDictionary=True the default property will not add nonexisting keys implicitly to prevent adding the key Missing when in Watch Window
' - if an item is set with a nonexisting index (.Count=3, Index=5) the storage is expanded to that index; all implicitly created items will be Empty
' - keys of items that have no keys will be Empty when using .Keys() or .KeyOfItemAt(); these empty keys can be omitted using .Keys(false)
'   Empty keys and items act like the default value of most datatypes: "" for String and 0 for numeric data types which simplifies for-each tasks
' - default property and default value are not to be confused:
'   1) the default property returns items for given indexes or keys
'   2) the default value (usually the first item) is returned when the default property is called without any argument or the argument Missing
'   .DefaultValueEnabled defines what happens when the default property is called without argument:
'   - True: returns either the first item, "[EMPTY]", or "[NONEXISTING]"
'   - False: returns a reference to the DictCollection itself or in some cases "[NONEXISTING]" when entering endless evaluation loops
'
' ----------- KNOWN BUGS ----------------------------------------------------------------------------------------------------------------------------
'   BUG001 (2018-09-08): With .EmulateDictionary=True, keys CInt(1) and CDate(1) are not recognizes as the same numeric key. The internal key
'       representation as String for Date and Numbers is different so it doesn't match. Converting Dates to Doubles introduces rounding problems.
'   BUG002 (2018-09-20): DictCollection Default Property Evaluation skips VB String concatenation with '&' or '+'. This code won't work:
'       1) Debug.Print "Test: " & dc("SomeNonexistingKey") will print "[NONEXISTING]" instead of "Test: [NONEXISTING]"
'       2) a = dc("KeyOfEmptyCollection") + "_" will store "[EMPTY]" in variable a instead of "[EMPTY]_"
'   BUG003 (2018-09-26): Expression 'If dc.AsDC(-1).CollectionType <> 0 Then ...' will execute Then clause although .CollectionType = 0
'       Workaround: use 'Set dc2 = dc.AsDC(-1): If dc2.CollectionType <> 0 Then ...' instead
'   BUG004 (2018-09-26): Numberic Default Values used in 'If ... Then ...' conditions do not evaluate correctly and will always execute the Then clause
'       Workaround: use 'Set dc2 = dc.AsDC(-1): If dc2.CollectionType <> 0 Then ...' instead
'
' ----------- KNOWN PROBLEMS ------------------------------------------------------------------------------------------------------------------------
' - With dc.DefaultValueEnabled=True, 'Is DictCollection' and 'VarType(dc)' do not work because default Property evaluation returns String.
'   Use TypeName() or '.DefaultValueEnabled=False'.
' - With .DefaultValueEnabled=True, setting the default value with Set keyword doesn't work: use 'dc = new Collection' instead of 'Set dc = new Collection'
' - With dc.DefaultValueEnabled=False, expressions like 'dc <> 0' lead into an endless default property evaluation loop that is stopped by force in the
'   code. This is done by using a default value evaluation counter and checking if it exceeds the number 10000.
' - With DefaultValueEnabled=False and ThrowErrors=False, default property will return "[NONEXISTING]" when called without argument.
'   This might be confused with the case when it is really nonexisting and DefaultValueEnabled=False
' - With dc.EmulateDictionary=True, monitoring a DictCollection object in the Watch Window while debugging code will sometimes add an
'   Empty item with Missing as key. This happens because the Watch Window is calling the default property .Item(Missing) and the behavior of
'   Scripting.Dictionary is to implicitly add all requested nonexisting keys (in this case Missing)
' - With dc.EmulateDictionary=True, numeric keys are not sorted correctly because they are stored as string, e.g. "[NUM:100]" comes before "[NUM:99]"
' - Existing keys of the form "_N_C" or "_N" will not be affected by .Flatten but .Unflatten will incorrectly remove these keys if N > dc.Count
' - When running .SelfTest and other test functions in Excel VBA Immediate Window, saving the file just before running the tests helps avoid crashes
'
' ----------- KNOWN DESIGN PROBLEMS -----------------------------------------------------------------------------------------------------------------
' - Default Value allows chainability but trade-off with VarType() not working is tricky and difficult to understand
' - Emulation of Collection introduces mExternalIndexOffset, making it complicated for private functions to use public index based functions
' - .AddDC() returns newly added subcollection while .Add() returns root collection which is inconsistent but useful
' - with .ThrowErrors=False existing keys are overwritten, there is no option to silently reject overwrites
'
' ----------- BREAKING CHANGES ----------------------------------------------------------------------------------------------------------------------
'   2018-10-08 V036     None (initial commit on GitHub)
'   2019-06-01 V040     Changed functions .CopyItems(), .Clone(), .Flatten(), .Unflatten() to support chainability
'       Old: CopyItems(Optional SourceCollection As DictCollection, Optional TargetCollection As DictCollection, Optional targetIndex As Long = -1) As DictCollection
'          New: CopyItems(TargetCollection As DictCollection, Optional targetIndex As Long = -1) As DictCollection
'           New: CopyItemsToNew() As DictCollection
'       Old: Flatten(Optional TargetCollection As DictCollection, Optional ReturnAnalysisResults As Boolean = False) As DictCollection
'           New: Flatten() As DictCollection
'           New: FlattenToNew() As DictCollection
'       Old: Unflatten(Optional TargetCollection As DictCollection)
'           New: Unflatten() As DictCollection
'           New: UnflattenToNew() as DictCollection
'       Old: Clone(Optional TargetCollection As DictCollection) As DictCollection
'           New: CloneToNew() As DictCollection
'           New: CloneTo(TargetCollection As DictCollection) As DictCollection
'           New: CloneToPreserve(TargetCollection As DictCollection) As DictCollection
'
' ----------- TODOS ---------------------------------------------------------------------------------------------------------------------------------
' - change .Flatten() and .Unflatten to support empty Arrays and Collections with "[EMPTY_ARRAY]" and "[EMPTY_OBJECT]"
' - change implementation of .MatchKeysToNew() to suppport wildcards
' - implement .FromPairs(Arr) -> creates new DictCollection from array containing key-value-pairs
' - implement .Swap(IndexA, IndexB) -> swaps two items
' - implement .AddKey(Key, Optional Index) -> adds a key; if the item at index already has a key, another key is being added for that item
' - implement .CollectionType = 6 -> keys only (will affect copy/clone)
' - implement .AddValuesFromRange(Values, IndexOfValues, Keys, IndexOfKeys, KeysAndValuesAreRows) -> adds values/keys from Excel Range
' - implement .AddValuesFromArray(Values, DimensionIndexOfValues, Keys, DimensionIndexOfKeys) -> adds values/keys from one/multi-dimensional arrays
' - implement .AddValuesFromDictCollection(AnotherDictCollection) -> adds the content from another DictCollection
' - implement .AsDouble(IndexOrKey, Optional DefaultValue) -> returns items as Double for typesafe usage
' - implement .AsLong(IndexOrKey, Optional DefaultValue) -> returns items as Long for typesafe usage
' - implement .AsDate(IndexOrKey, Optional DefaultValue) -> returns items as Date for typesafe usage
' - implement .AsString(IndexOrKey, Optional DefaultValue) -> returns items as String for typesafe usage
' - set empty/missing item and key values as private constants for better hackability
'
' ======================== END: DOCUMENTATION =======================================================================================================

Option Explicit

Private Const THROW_ERRORS_DEFAULT As Boolean = False
Private Const COMPAREMODE_DEFAULT As Integer = VbCompareMethod.vbBinaryCompare ' Compatible with Scripting.Dictionary
Private Const LAZYSORTING_DEFAULT As Boolean = True
Private Const DEFAULTVALUEENABLED_DEFAULT As Boolean = True
Private Const ZEROBASEDINDEX_DEFAULT As Boolean = True

' returned when ThrowErrors=False and acccessing a nonexisting key or index with default value: Set d = new DictCollection: Debug.Print d(0) -> "[NONEXISTING]"
Private Const NONEXISTING_VALUE_DEFAULT = "[NONEXISTING]"
' returned when ThrowErrors=False and DefaultValueEnabled=True and acccessing an empty collection with default value: Set d = new DictCollection: Debug.Print d -> "[EMPTY]"
Private Const EMPTY_COLLECTION_VALUE_DEFAULT = "[EMPTY]"

' "all" prefixes to put variables on top of property list in watch window when debugging
Private allItems() As Variant, allItemKeyIndexes() As Long
Private allKeys() As String, allKeyItemIndexes() As Long
Private mThrowErrors As Boolean ' true: will throw errors when accessing nonexisting items or using invalid types
Private mCompareMode As VbCompareMethod ' switches between case sensitive and case insensitive key comparison
' vbBinaryCompare = 0   default case sensitive comparison, sort order is derived from internal binary representation,
'                       unix-filename-like sorting (uppercase chars are sorted before lowercase chars)
' vbTextCompare = 1     case insensitive comparison, based on the order in the ASCII table
'                       will ignore all local character derivations: · = A = a = ‰ = ƒ
' vbDatabaseCompare = 2 case insensitive, can only be used in Microsoft Access VBA
'                       will determine the sort order using the local ID of the database
Private mLazySorting As Boolean ' true: will not sort keys while adding but later when keys are searched
Private mKeysAreSortedAndUnique As Boolean ' true: allKeys array is sorted and contains no temporary key duplicates
Private mExternalIndexOffset As Long ' 1: Collection compatible, 0: Scripting.Dictionary compatible
Private mEmulateCollection As Boolean ' if true, emulates Collection behavior
Private mEmulateScriptingDictionary As Boolean ' if true, emulates Scripting.Dictionary behavior
Private mDefaultValueEnabled As Boolean ' if true, returns the first item as default value or if empty "[EMPTY]" and if nonexisting "[NONEXISTING"]
Private mNonExistingValue As Variant ' returned when ThrowErrors=False and acccessing a nonexisting key or index with default value: Set d = new DictCollection: Debug.Print d(0)
Private mEmptyCollectionValue As Variant ' returned when ThrowErrors=False and acccessing an empty collection with default value: Set d = new DictCollection: Debug.Print d
Private mItemCount As Long ' number of items. if item array is uninitialized (nonexisting state), then number of items = -1
Private mKeyCount As Long ' number of keys. if keys array is uninitialized (array state), then number of keys = -1
Private mScriptingDictionaryObjectKeys As Variant ' used by Scripting.Dictionary compatibility: holds object references where objects are used as keys
Private mScriptingDictionaryStoredKeyTypes As DictCollection ' used by Scripting.Dictionary compatibility: holds the key types of number keys (Integer, Long, Single, Double, Currency, Decimal, Date)
Private mDefaultPropertyWithMissingParamCounter As Long ' used to avoid endless default property evaluation loops

Private Type FindKeyResult
    Exact As Boolean
    Index As Long
End Type
Private Type EvaluateIndexOrKeyResult
    WasIndex As Boolean
    WasKey As Boolean
    WasFound As Boolean
    IsObject As Boolean
    Index As Long
End Type
Private Type EvaluateKeyResult
    WasNumber As Boolean
    WasString As Boolean
    WasEmpty As Boolean
    WasInvalid As Boolean
    key As String
End Type
Private Type WalkDictCollectionTreeOptions
    Analyze As Boolean
    Clone As Boolean
    Flatten As Boolean
    Unflatten As Boolean
    EnsureKeys As Boolean
    Stringify As Boolean
End Type


' ======================== START: INITIALIZATION AND STATES =========================================================================================
Private Sub Class_Initialize(): init: End Sub
Private Sub Class_Terminate(): End Sub ' freeing used resources manually in this event causes excel to crash when using this a lot
Private Sub init()
    applyDefaultAccessSettings
    mLazySorting = LAZYSORTING_DEFAULT
    mKeysAreSortedAndUnique = False
    initializeItemArray
    uninitializeKeyArray
End Sub
Private Sub applyDefaultAccessSettings()
    mThrowErrors = THROW_ERRORS_DEFAULT
    mCompareMode = COMPAREMODE_DEFAULT
    mNonExistingValue = NONEXISTING_VALUE_DEFAULT
    mEmptyCollectionValue = EMPTY_COLLECTION_VALUE_DEFAULT
    mExternalIndexOffset = 0
    mDefaultValueEnabled = DEFAULTVALUEENABLED_DEFAULT
End Sub
Private Sub initializeItemArray(): allItems = VBA.Array(): Erase allItemKeyIndexes: mItemCount = 0: End Sub
Private Sub initializeKeyArray(): allKeys = Split("", "."): Erase allKeyItemIndexes: mKeyCount = 0: End Sub
Private Sub initializeDictionaryObjectKeyArray(): mScriptingDictionaryObjectKeys = VBA.Array(): End Sub
Private Sub initializeDictionaryStoredKeyTypes(): Set mScriptingDictionaryStoredKeyTypes = New DictCollection: mScriptingDictionaryStoredKeyTypes.NonExistingValue = "": End Sub
Private Sub uninitializeKeyArray(): Erase allKeys: Erase allKeyItemIndexes: mKeyCount = -1: End Sub
Private Sub uninitializeItemArray(): Erase allItems: Erase allItemKeyIndexes: mItemCount = -1: End Sub
Private Sub uninitializeDictionaryObjectKeyArray(): mScriptingDictionaryObjectKeys = Empty: End Sub
Private Sub uninitializeDictionaryStoredKeyTypes(): Set mScriptingDictionaryStoredKeyTypes = Nothing: End Sub

Public Property Get CollectionTypeNonexisting() As Integer: CollectionTypeNonexisting = 0: End Function
Public Property Get CollectionTypeEmptyArray() As Integer: CollectionTypeEmptyArray = 1: End Function
Public Property Get CollectionTypeFilledArray() As Integer: CollectionTypeFilledArray = 2: End Function
Public Property Get CollectionTypeEmptyKeyValueStore() As Integer: CollectionTypeEmptyKeyValueStore = 3: End Function
Public Property Get CollectionTypeFilledKeyValueStore() As Integer: CollectionTypeFilledKeyValueStore = 4: End Function
Public Property Get CollectionTypeFilledKeyValueStoreWithSomeItemsHavingNoKey() As Integer: CollectionTypeFilledKeyValueStoreWithSomeItemsHavingNoKey = 5: End Function

' 0 = nonexisting, 1 = empty array, 2 = filled array, 3 = empty key-value-store, 4 = filled key-value-store, 5 = filled key-value-store with some items having no key
Public Property Get CollectionType() As Integer
    If mItemCount < 0 Then
        CollectionType = 0 ' nonexisting
    ElseIf mItemCount = 0 Then
        If mKeyCount < 0 Then
            CollectionType = 1 ' empty array
        Else
            CollectionType = 3 ' empty key-value-store
        End If
    Else
        If mKeyCount < 0 Then
            CollectionType = 2 ' filled array
        Else
            Dim i As Long
            For i = 0 To (mItemCount - 1)
                If allItemKeyIndexes(i) = -1 Then
                    CollectionType = 5 ' filled key-value-store with some items having no key
                    Exit Property
                End If
            Next
            CollectionType = 4 ' filled key-value-store
        End If
    End If
End Property
Public Property Let CollectionType(Value As Integer)
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
        Case 3 ' empty key-value-store
            initializeItemArray
            initializeKeyArray
        Case 4 ' filled key-value-store
            If mItemCount < 0 Then initializeItemArray
            If mKeyCount < 0 Then initializeKeyArray
            EnsureAllItemsHaveKeys
        Case 5 ' filled key-value-store with some items having no key
            If mItemCount < 0 Then initializeItemArray
            If mKeyCount < 0 Then initializeKeyArray
    End Select
End Property
' sets all item-to-key pointer to -1 (none), needed to remove all keys
Private Sub setAllItemKeyIndexesToNone()
    Dim i As Long
    For i = 0 To (mItemCount - 1): allItemKeyIndexes(i) = -1: Next
End Sub
Private Function increaseItemArray() As Long
    increaseItemArray = UBound(allItems) + 1
    ReDim Preserve allItems(0 To increaseItemArray)
    ReDim Preserve allItemKeyIndexes(0 To increaseItemArray)
    mItemCount = increaseItemArray + 1
End Function
Private Function increaseItemArrayTo(newUpperBound As Long)
    Dim oldUpperBound As Long, Index As Long
    oldUpperBound = UBound(allItems)
    If newUpperBound > oldUpperBound Then
        ReDim Preserve allItems(0 To newUpperBound)
        ReDim Preserve allItemKeyIndexes(0 To newUpperBound)
        For Index = oldUpperBound + 1 To newUpperBound
            allItemKeyIndexes(Index) = -1 ' set all new key indexes to -1
        Next
        mItemCount = newUpperBound + 1
    End If
End Function
Private Function decreaseItemArray() As Long
    decreaseItemArray = UBound(allItems) - 1
    If decreaseItemArray < 0 Then
        initializeItemArray
    Else
        ReDim Preserve allItems(decreaseItemArray)
        ReDim Preserve allItemKeyIndexes(decreaseItemArray)
        mItemCount = decreaseItemArray + 1
    End If
End Function
Private Function increaseKeyArray() As Long
    increaseKeyArray = UBound(allKeys) + 1
    ReDim Preserve allKeys(0 To increaseKeyArray)
    ReDim Preserve allKeyItemIndexes(0 To increaseKeyArray)
    mKeyCount = increaseKeyArray + 1
End Function
Private Function decreaseKeyArray() As Long
    decreaseKeyArray = UBound(allKeys) - 1
    If decreaseKeyArray < 0 Then
        uninitializeKeyArray
    Else
        ReDim Preserve allKeys(decreaseKeyArray)
        ReDim Preserve allKeyItemIndexes(decreaseKeyArray)
        mKeyCount = decreaseKeyArray + 1
    End If
End Function
Public Property Get EmptyCollectionValue(): EmptyCollectionValue = mEmptyCollectionValue: End Property
Public Property Let EmptyCollectionValue(Value As Variant): mEmptyCollectionValue = Value: End Property
Public Property Get NonExistingValue(): NonExistingValue = mNonExistingValue: End Property
Public Property Let NonExistingValue(Value As Variant): mNonExistingValue = Value: End Property
Public Property Get CompareMode() As VbCompareMethod: CompareMode = mCompareMode: End Property
Public Property Let CompareMode(Value As VbCompareMethod)
    If mCompareMode <> Value Then
        If mKeyCount > 0 And mThrowErrors Then
            ' Emulate Scripting.Dictionary behavior where change of CompareMode is only possible when Dictionary is empty
            Err.Raise 5, "Invalid procedure call or argument": Exit Property
        Else
            mCompareMode = Value
            If mKeyCount > 0 Then sortKeys 0, (mKeyCount - 1) 're-sort keys so that internalFindKeyIndex will work as intended
        End If
    End If
End Property
Public Property Get ThrowErrors() As Boolean: ThrowErrors = mThrowErrors: End Property
Public Property Let ThrowErrors(Value As Boolean): mThrowErrors = Value: End Property
Public Property Get LazySorting() As Boolean: LazySorting = mLazySorting: End Property
Public Property Let LazySorting(Value As Boolean)
    mLazySorting = Value
    ' sort if necessary
    If Not mLazySorting And mKeyCount > 1 And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
End Property
Public Property Get ZeroBasedIndex() As Boolean: ZeroBasedIndex = (mExternalIndexOffset = 0): End Property
Public Property Let ZeroBasedIndex(Value As Boolean)
    If Value Then mExternalIndexOffset = 0 Else mExternalIndexOffset = 1
End Property
Public Property Get DefaultValueEnabled() As Boolean: DefaultValueEnabled = mDefaultValueEnabled: End Property
Public Property Let DefaultValueEnabled(Value As Boolean): mDefaultValueEnabled = Value: End Property
Public Property Get EmulateCollection() As Boolean:  EmulateCollection = mEmulateCollection: End Property
Public Property Let EmulateCollection(Value As Boolean)
    mEmulateCollection = Value
    If mEmulateCollection Then
        mCompareMode = vbTextCompare
        mExternalIndexOffset = 1
        mDefaultValueEnabled = False
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
Public Property Get EmulateDictionary() As Boolean:  EmulateDictionary = mEmulateScriptingDictionary: End Property
Public Property Let EmulateDictionary(Value As Boolean)
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

' copies all settings from one DictCollection to another that allows endless subcollection chaining with modified NonExistingValue and EmptyCollectionValue to work properly
Public Sub CopySubCollectionChainingSettings(FromCollection As DictCollection, ToCollection As DictCollection)
    If (FromCollection Is Nothing) Or (ToCollection Is Nothing) Then Exit Sub
    ToCollection.ThrowErrors = FromCollection.ThrowErrors
    ToCollection.NonExistingValue = FromCollection.NonExistingValue
    ToCollection.EmptyCollectionValue = FromCollection.EmptyCollectionValue
End Sub

' copies all settings from one DictCollection to another
Public Sub CopyAllSettingsExceptCollectionType(FromCollection As DictCollection, ToCollection As DictCollection)
    If (FromCollection Is Nothing) Or (ToCollection Is Nothing) Then Exit Sub
    ' copy these two first as these also implicitly set .ThrowErrors, .DefaultValueEnabled and .NonExistingValue
    ' only apply them if they are different to prevent loosing Dictionary original key types
    If ToCollection.EmulateCollection <> FromCollection.EmulateCollection Then ToCollection.EmulateCollection = FromCollection.EmulateCollection
    If ToCollection.EmulateDictionary <> FromCollection.EmulateDictionary Then ToCollection.EmulateDictionary = FromCollection.EmulateDictionary
    ' copy all others
    ToCollection.NonExistingValue = FromCollection.NonExistingValue
    ToCollection.EmptyCollectionValue = FromCollection.EmptyCollectionValue
    ToCollection.ZeroBasedIndex = FromCollection.ZeroBasedIndex
    ToCollection.CompareMode = FromCollection.CompareMode
    ToCollection.LazySorting = FromCollection.LazySorting
    ToCollection.DefaultValueEnabled = FromCollection.DefaultValueEnabled
    ToCollection.ThrowErrors = FromCollection.ThrowErrors
End Sub
' ======================== END: INITIALIZATION AND STATES ===========================================================================================


' ======================== START: ITEM, ITEMAT, ITEMOF, ADD, INSERT, ADD2, EXISTS, REMOVE, REMOVEALL ================================================
' Get item by index or key. If ThrowErrors = False (default), then
' 1) nonexisting indexes/keys will return placeholder string "[NONEXISTING]"
'    this allows endless subcollection chaining with nonexisting indexes/keys:
'    example: dc("existing")("existing")("nonexisting")("nonexisting") will result in "[NONEXISTING]"
' 2) empty nested dictcollections will return placeholder string "[EMPTY]"
'    example: dc("existing")("empty") will result in "[EMPTY]"
Public Property Get Item(IndexOrKey As Variant) As Variant
    Dim result As EvaluateIndexOrKeyResult
    result = internalGetItem(IndexOrKey, mEmulateScriptingDictionary) ' implicitly add nonexisting keys when emulating Scripting.Dictionary
    If result.WasFound Then
        If result.IsObject Then Set Item = allItems(result.Index) Else Item = allItems(result.Index)
    Else
        If mDefaultValueEnabled Then
            Dim newDictCollection As DictCollection: Set newDictCollection = New DictCollection
            CopySubCollectionChainingSettings Me, newDictCollection
            newDictCollection.CollectionType = 0 ' set to nonexisting
            newDictCollection.ThrowErrors = False
            Set Item = newDictCollection
        Else
            Item = mNonExistingValue
        End If
    End If
End Property

' Sets an item to basic type. If ThrowErrors=True and item does not exist, throws errors.
' If ThrowErrors=False (Default), will add the item by given index or key
Public Property Let Item(Optional IndexOrKey As Variant, Value As Variant)
    If mEmulateScriptingDictionary Then
        Dim key2 As String: key2 = convertScriptingDictionaryToDictCollectionKey(IndexOrKey)
        internalSetItem key2, Value
        storeScriptingDictionaryKeyInformation IndexOrKey, key2
    ElseIf mEmulateCollection Then
        internalSetItem convertCollectionToDictCollectionKey(IndexOrKey), Value
    Else
        internalSetItem IndexOrKey, Value
    End If
End Property

' same as Property Let
Public Property Set Item(Optional IndexOrKey As Variant, Value As Variant): Me.Item(IndexOrKey) = Value: End Property

' same as Property Let but chainable
Public Function SetItem(IndexOrKey As Variant, Value As Variant) As DictCollection: Me.Item(IndexOrKey) = Value: Set SetItem = Me: End Function

' evaluates the index or key, searches for an existing item, throws errors if needed and returns a result with the following information
' 1) was the item found 2) was it by index or key 3) the index of the found item (-1 if none) and 4) is the item an object
Private Function internalGetItem(IndexOrKey As Variant, Optional ImplicitlyAddNonexistingKeys As Boolean = False) As EvaluateIndexOrKeyResult
    Dim retVal As EvaluateIndexOrKeyResult
    If IsMissing(IndexOrKey) And Not mEmulateScriptingDictionary Then
        ' Emulate Collection behavior
        If mThrowErrors Then Err.Raise 5, "DictCollection", "Invalid procedure call or argument"
    ElseIf mItemCount > 0 Then
        ' sort keys if needed
        If mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
        ' evaluate index or key
        If mEmulateScriptingDictionary Then
            retVal = evaluateIndexOrKey(convertScriptingDictionaryToDictCollectionKey(IndexOrKey))
        ElseIf mEmulateCollection Then
            If TypeName(IndexOrKey) = "Nothing" Then
                Err.Raise 13, "DictCollection", "Type mismatch"
            Else
                retVal = evaluateIndexOrKey(convertCollectionToDictCollectionKey(IndexOrKey))
            End If
        Else
            retVal = evaluateIndexOrKey(IndexOrKey)
        End If
    End If
    If retVal.WasFound Then
        retVal.IsObject = IsObject(allItems(retVal.Index))
    Else
        ' key was not found
        If mThrowErrors And Not mEmulateScriptingDictionary Then
            If retVal.WasIndex Then
                ' Emulate Collection behavior
                If mItemCount < 1 Then
                    Err.Raise 5, "DictCollection", "Invalid procedure call or argument"
                Else
                    Err.Raise 9, "DictCollection", "Subscript out of range"
                End If
            ElseIf retVal.WasKey = False Then
                ' was neither string nor number, emulate Collection behavior
                If mItemCount < 1 Then
                    Err.Raise 5, "DictCollection", "Invalid procedure call or argument"
                Else
                    If IsEmpty(IndexOrKey) Then
                        Err.Raise 9, "DictCollection", "Subscript out of range"
                    Else
                        Err.Raise 13, "DictCollection", "Type mismatch"
                    End If
                End If
            Else
                Err.Raise 5, "DictCollection", "Invalid procedure call or argument"
            End If
        End If
        ' Emulate Scripting.Dictionary behavior of implicitly adding nonexisting keys with empty items when they are accessed
        If ImplicitlyAddNonexistingKeys Then Me.Add IndexOrKey, Empty
    End If
    internalGetItem = retVal
End Function

' Alternative (and fast) property for explicitly accessing an item by index (for instance when emulating Scripting.Dictionary and index numbers are interpreted as keys)
Public Property Get ItemAt(Index As Long) As Variant
    Dim index2 As Long: index2 = Index - mExternalIndexOffset
    ItemAt = mNonExistingValue ' default value for not found
    If mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
    If index2 < 0 Or index2 >= mItemCount Then
        ' Emulate Collection behavior
        If mThrowErrors Then Err.Raise 9, "DictCollection", "Subscript out of range": Exit Property
    Else
        If IsObject(allItems(index2)) Then Set ItemAt = allItems(index2) Else ItemAt = allItems(index2)
    End If
End Property
' Alternative (and fast) property for explicitly setting an item by index (for instance when emulating Scripting.Dictionary and index numbers are interpreted as keys)
Public Property Let ItemAt(Index As Long, Value As Variant)
    Dim index2 As Long: index2 = Index - mExternalIndexOffset
    If index2 < 0 Then
        ' Emulate Collection behavior
        If mThrowErrors Then Err.Raise 9, "DictCollection", "Subscript out of range"
        Exit Property
    End If
    If mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
    If index2 >= mItemCount Then
        ' Emulate Collection behavior
        If mThrowErrors Then Err.Raise 9, "DictCollection", "Subscript out of range"
        increaseItemArrayTo index2
    End If
    If IsObject(Value) Then Set allItems(index2) = Value Else allItems(index2) = Value
End Property
' same as Property Let
Public Property Set ItemAt(Index As Long, Value As Variant): ItemAt(Index) = Value: End Property

' Alternative (and fast) property for explicitly accessing an item by key
Public Property Get ItemOf(key As String) As Variant
    ItemOf = mNonExistingValue ' set to nonexisting
    If mKeyCount < 1 Or mItemCount < 1 Then
        If mThrowErrors Then Err.Raise 5, "DictCollection", "Key does not exist"
        Exit Property
    End If
    If mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
    If Len(key) = 0 Then
        If Not mEmulateCollection And Not mEmulateScriptingDictionary Then
            If mThrowErrors Then Err.Raise 5, "DictCollection", "Empty keys are not supported"
            Exit Property
        End If
    End If
    Dim foundKey As FindKeyResult
    ' sort keys if needed
    If mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
    If mEmulateScriptingDictionary Then
        foundKey = internalFindKeyIndex(convertScriptingDictionaryToDictCollectionKey(key), mCompareMode)
    ElseIf mEmulateCollection Then
        foundKey = internalFindKeyIndex(convertCollectionToDictCollectionKey(key), mCompareMode)
    Else
        foundKey = internalFindKeyIndex(key, mCompareMode)
    End If
    If foundKey.Exact Then
        Dim ItemIndex: ItemIndex = allKeyItemIndexes(foundKey.Index)
        If IsObject(allItems(ItemIndex)) Then Set ItemOf = allItems(ItemIndex) Else ItemOf = allItems(ItemIndex)
    Else
        If mThrowErrors Then Err.Raise 5, "DictCollection", "Key does not exist"
    End If
End Property

' Alternative (and fast) property for explicitly setting an item by key
Public Property Let ItemOf(key As String, Value As Variant)
    Dim foundKey As FindKeyResult
    If Len(key) = 0 Then
        If Not mEmulateCollection And Not mEmulateScriptingDictionary Then
            If mThrowErrors Then Err.Raise 5, "DictCollection", "Empty keys are not supported"
            Exit Property
        End If
    End If
    ' sort keys if needed
    If mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
    If mEmulateScriptingDictionary Then
        foundKey = internalFindKeyIndex(convertScriptingDictionaryToDictCollectionKey(key), mCompareMode)
    ElseIf mEmulateCollection Then
        foundKey = internalFindKeyIndex(convertCollectionToDictCollectionKey(key), mCompareMode)
    Else
        foundKey = internalFindKeyIndex(key, mCompareMode)
    End If
    If foundKey.Exact Then
        ' found, set to value
        Dim ItemIndex: ItemIndex = allKeyItemIndexes(foundKey.Index)
        If IsObject(Value) Then Set allItems(ItemIndex) = Value Else allItems(ItemIndex) = Value
        If StrComp(key, allKeys(foundKey.Index), vbBinaryCompare) <> 0 Then allKeys(foundKey.Index) = key ' keys are not exactly the same, overwrite with given key
    Else
        If mThrowErrors Then Err.Raise 5, "DictCollection", "Key does not exist"
        If key <> "" Then
            Dim newItemIndex As Long, newKeyIndex As Long
            If mItemCount = -1 Then initializeItemArray
            increaseItemArray
            newItemIndex = mItemCount - 1
            If IsObject(Value) Then Set allItems(newItemIndex) = Value Else allItems(newItemIndex) = Value
            If mKeyCount = -1 Then initializeKeyArray
            If mLazySorting Then newKeyIndex = mKeyCount Else newKeyIndex = foundKey.Index
            allItemKeyIndexes(newItemIndex) = newKeyIndex
            insertKey key, newKeyIndex, newItemIndex
            If mLazySorting Then mKeysAreSortedAndUnique = False
        Else
            allItemKeyIndexes(newItemIndex) = -1 ' set to -1 = no key exists for this item
        End If
    End If
End Property
' same as Property Let
Public Property Set ItemOf(key As String, Value As Variant): ItemOf(key) = Value: End Property

' The default property of the DictCollection class. Allows expressions like this:
'   'item = dc(0)' or 'item = dc("keyA")' or 'defaultValue = dc()' or 'If dc <> "[NONEXISTING]" Then ...' or 'If dc <> "[EMPTY]" Then ...'
' Returns the default value (first item or [EMPTY] or [NONEXISTING] when called without argument and .DefaultValueEnabled=True
' The values of [EMPTY] and [NONEXISTING] can be modified using the properties .EmptyCollectionValue and .NonExistingValue
' Identical to .Item() when called with an argument with the exception that it does not implicitly add nonexisting keys when .EmulateDictionary=True
Public Property Get DefaultProperty(Optional IndexOrKey As Variant) As Variant
Attribute DefaultProperty.VB_UserMemId = 0
    ' return a reference to this class instance or the default value when called without argument or with Missing als argument
    If IsMissing(IndexOrKey) And Not mEmulateScriptingDictionary Then
        If mDefaultValueEnabled Then
            If mItemCount = -1 Then DefaultProperty = mNonExistingValue: Exit Property  ' nonexisting value
            If mItemCount = 0 Then
                DefaultProperty = mEmptyCollectionValue ' value of empty DictCollection
            Else
                If IsObject(allItems(0)) Then Set DefaultProperty = allItems(0) Else DefaultProperty = allItems(0)
            End If
        Else
            If mDefaultPropertyWithMissingParamCounter > 10000 Then
                ' safety exit to avoid endless default property evaluation loops
                DefaultProperty = mNonExistingValue
                mDefaultPropertyWithMissingParamCounter = 0
            Else
                Set DefaultProperty = Me ' set to object if default value is disabled and not in default value evaluation loop
                mDefaultPropertyWithMissingParamCounter = mDefaultPropertyWithMissingParamCounter + 1
            End If
        End If
        Exit Property
    End If
    
    Dim result As EvaluateIndexOrKeyResult
    result = internalGetItem(IndexOrKey, False)
    If result.WasFound Then
        If result.IsObject Then Set DefaultProperty = allItems(result.Index) Else DefaultProperty = allItems(result.Index)
    Else
        If mDefaultValueEnabled Then
            Dim newDictCollection As DictCollection: Set newDictCollection = New DictCollection
            CopySubCollectionChainingSettings Me, newDictCollection
            newDictCollection.CollectionType = 0 ' set to nonexisting
            newDictCollection.ThrowErrors = False
            Set DefaultProperty = newDictCollection
        Else
            DefaultProperty = mNonExistingValue
        End If
    End If
mDefaultPropertyWithMissingParamCounter = 0
End Property

' The default property of the DictCollection class. Assigns an item to an index or key like this: 'dc(3) = "Hello"' or 'dc("keyA") = True'
' Sets the default value when used without argument and DefaultValueEnabled=True: 'dc() = 1.5' and 'dc = 1.5' do the same as 'dc.Item(0) = 1.5'
Public Property Let DefaultProperty(Optional IndexOrKey As Variant, Value As Variant)
    If IsMissing(IndexOrKey) And Not mDefaultValueEnabled Then
        If mThrowErrors Then Err.Raise 5, "Cannot assign default value with DefaultValueEnabled=False"
        Exit Property ' do nothing
    End If
    ' accept Missing as key if .EmulateDictionary=True or set Index to 0 or 1 = first item
    If IsMissing(IndexOrKey) And Not mEmulateScriptingDictionary Then IndexOrKey = mExternalIndexOffset
    If mEmulateScriptingDictionary Then
        Dim key2 As String: key2 = convertScriptingDictionaryToDictCollectionKey(IndexOrKey)
        internalSetItem key2, Value
        storeScriptingDictionaryKeyInformation IndexOrKey, key2
    ElseIf mEmulateCollection Then
        internalSetItem convertCollectionToDictCollectionKey(IndexOrKey), Value
    Else
        internalSetItem IndexOrKey, Value
    End If
End Property

' The default property of the DictCollection class. Assigns an object to an index or key like this: 'dc(3) = "Hello"' or 'dc("keyA") = True'
' Sets the default value when used without argument and DefaultValueEnabled=True: 'dc() = MyObject' and 'dc = MyObject' do the same as 'Set dc.Item(0) = MyObject'
Public Property Set DefaultProperty(Optional IndexOrKey As Variant, Value As Variant)
    If IsMissing(IndexOrKey) And Not mDefaultValueEnabled Then
        If mThrowErrors Then Err.Raise 5, "Cannot assign default value with DefaultValueEnabled=False"
        Exit Property ' do nothing
    End If
    ' accept Missing as key if .EmulateDictionary=True or set Index to 0 or 1 = first item
    If IsMissing(IndexOrKey) And Not mEmulateScriptingDictionary Then IndexOrKey = mExternalIndexOffset
    If mEmulateScriptingDictionary Then
        Dim key2 As String: key2 = convertScriptingDictionaryToDictCollectionKey(IndexOrKey)
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
Private Sub internalSetItem(Optional IndexOrKey As Variant, Optional Value As Variant)
    Dim result As EvaluateIndexOrKeyResult, action As Integer ' 0 = found, 1 = not found/byIndex, 2 = not found/byKey, 3 = not found/default value
    If IsMissing(IndexOrKey) Then
        If mItemCount > 0 Then
            ' set default item (first item)
            result.Index = 0: result.WasIndex = True
            action = 0 ' found
        Else
            action = 3 ' not found/default value
        End If
    Else
        If mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
        result = evaluateIndexOrKey(IndexOrKey)
        If result.Index < 0 Or result.Index >= mItemCount Then
            ' key or index was not found
            If result.WasIndex Then
                ' index was found
                If mThrowErrors Then
                    ' Emulate Collection/Array behavior
                    Err.Raise 9, "DictCollection", "Subscript out of range"
                    Exit Sub
                End If
                If result.Index < 0 Then Exit Sub ' do nothing of negative index is passed
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
                ' was neither string nor number
                If mThrowErrors Then
                    If mEmulateScriptingDictionary Then
                        Err.Raise 91, "DictCollection", "Object variable or With block variable not set"
                    Else
                        Err.Raise 13, "DictCollection", "Keys of type" & TypeName(IndexOrKey) & " are not supported"
                    End If
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
                ' if object is another DictCollection, copy default value settings to ensure that subcollection chaining works globally the same
                If TypeName(Value) = "DictCollection" Then ' expression "Value Is DictCollection" cannot be used because of default value evaluation
                    Dim dc2 As DictCollection: Set dc2 = Value
                    CopySubCollectionChainingSettings Me, dc2
                End If
            Else
                allItems(result.Index) = Value
            End If
            ' check if item has a key and overwrite key if needed
            If allItemKeyIndexes(result.Index) > -1 And result.WasKey Then
                Dim currentKey As String: currentKey = allKeys(allItemKeyIndexes(result.Index))
                If StrComp(currentKey, IndexOrKey, vbBinaryCompare) <> 0 Then
                    ' new key is different (for instance with case insensitive text compare), overwrite old key with new key
                    allKeys(allItemKeyIndexes(result.Index)) = IndexOrKey
                    If mEmulateScriptingDictionary Then
                        ' also update the key in the type storage
                        If isScriptingDictionaryStoredKeyType(currentKey) Then mScriptingDictionaryStoredKeyTypes.key(currentKey) = IndexOrKey
                    End If
                End If
            End If
        Case 1: ' not found/byIndex
            internalAddItem Value, "", CLng(IndexOrKey)
        Case 2: ' not found/byKey
            internalAddItem Value, CStr(IndexOrKey)
        Case 3: ' not found/default value
            internalAddItem Value, "", 0
    End Select
End Sub

' Compatible with Dictionary .Add function; chainable
Public Function Add(Optional key As Variant = "", Optional Item As Variant) As DictCollection
    Dim result As EvaluateKeyResult
    Set Add = Me
    If IsMissing(Item) Then
        ' Set to empty or Emulate Scripting.Dictionary
        If Not mThrowErrors Then Item = Empty Else Err.Raise 450, "DictCollection", "Wrong number of arguments or invalid property assignement"
    End If
    If mEmulateScriptingDictionary Then
        result.key = convertScriptingDictionaryToDictCollectionKey(key)
        result.WasInvalid = False: result.WasString = True ' explicitly set to String to bypass error handling below
    ElseIf mEmulateCollection Then
        result = evaluateKey(convertCollectionToDictCollectionKey(key))
    Else
        result = evaluateKey(key)
    End If
    If result.WasInvalid Then
        If mThrowErrors Then Err.Raise 13, "DictCollection", "Keys of type " & TypeName(key) & " are not supported"
        Exit Function
    End If
    internalAddItem Item, result.key
    If mEmulateScriptingDictionary Then storeScriptingDictionaryKeyInformation key, result.key
End Function

' Inserts an item at a given index; chainable
' AtIndex = -1 (default) will insert the item after the last used index
' If the given key value is already used, the item previously associated with that key will be removed
' In that case, if the removed item's index was less than the newly added items index ...
Public Function Insert(Optional Item As Variant = Empty, Optional AtIndex As Long = -1, Optional key As Variant) As DictCollection
    Dim result As EvaluateKeyResult
    If (AtIndex <> -1) And (AtIndex - mExternalIndexOffset) < 0 Then
        ' Emulate Collection behavior for all index < 1 except -1 (insert to end)
        If mThrowErrors Then Err.Raise 9, "DictCollection", "Subscript out of range"
        Exit Function
    End If
    If mEmulateScriptingDictionary Then
        result.key = convertScriptingDictionaryToDictCollectionKey(key)
        result.WasInvalid = False: result.WasString = True ' explicitly set to string
    ElseIf mEmulateCollection Then
        ' Emulate Collection behavior of not accepting keys other than String or Missing (no key)
        If TypeName(key) <> "String" And Not IsMissing(key) Then
            If mThrowErrors Then Err.Raise 13, "Type mismatch"
            Exit Function
        End If
        result = evaluateKey(convertCollectionToDictCollectionKey(key))
    Else
        If IsMissing(key) Then key = ""
        result = evaluateKey(key)
    End If
    If result.WasInvalid Then
        If mThrowErrors Then Err.Raise 13, "DictCollection", "Keys of type " & TypeName(key) & " are not supported"
    End If
    internalAddItem Item, result.key, AtIndex ' will substract mExternalIndexOffset inside
    If mEmulateScriptingDictionary Then storeScriptingDictionaryKeyInformation key, result.key
    Set Insert = Me
End Function

' Compatible with Collection .Add function; chainable
Public Function Add2(Item As Variant, Optional key As Variant, Optional Before, Optional after) As DictCollection
    Dim AtIndex As Long, result As EvaluateKeyResult, tempIndex As Long, tname As String
    If IsMissing(Before) And IsMissing(after) Then
        AtIndex = -1 ' undefined
    Else
        ' Emulate Collection error behavior
        If IsObject(Before) Then tname = "Object" Else tname = TypeName(Before)
        Select Case tname
            Case "String": If mThrowErrors Then Err.Raise 5, "DictCollection", "Invalid procedure call or argument"
            Case "Empty": If mThrowErrors Then Err.Raise 9, "DictCollection", "Subscript out of range"
            Case "Integer", "Long", "Single", "Double", "Currency", "Decimal", "Date":
                tempIndex = Before - mExternalIndexOffset - 1
                If (tempIndex < 0 Or tempIndex >= (mItemCount - 1)) Then
                    If mThrowErrors Then Err.Raise 9, "DictCollection", "Subscript out of range"
                    If tempIndex < 0 Then GoTo ExitFunction
                End If
                AtIndex = Before - 1 ' offset will be substracted later in internalSetItem
            Case Else:
                If mThrowErrors And Not IsMissing(Before) Then Err.Raise 13, "DictCollection", "Type mismatch"
        End Select
        If IsObject(after) Then tname = "Object" Else tname = TypeName(after)
        Select Case tname
            Case "String": If mThrowErrors Then Err.Raise 5, "DictCollection", "Invalid procedure call or argument"
            Case "Empty": If mThrowErrors Then Err.Raise 9, "DictCollection", "Subscript out of range"
            Case "Integer", "Long", "Single", "Double", "Currency", "Decimal", "Date":
                tempIndex = after - mExternalIndexOffset + 1
                ' Emulate Collection behavior
                If (tempIndex <= 0 Or tempIndex > mItemCount) Then
                    If mThrowErrors Then Err.Raise 9, "DictCollection", "Subscript out of range"
                    If tempIndex < 0 Then GoTo ExitFunction
                End If
                AtIndex = after + 1 ' offset will be substracted later in internalSetItem
            Case Else:
                If mThrowErrors And Not IsMissing(after) Then Err.Raise 13, "DictCollection", "Type mismatch"
        End Select
    End If
    If mThrowErrors Then
        ' Emulate Collection error behavior
        If IsNumeric(key) Or IsDate(key) Then
            Dim result2 As EvaluateIndexOrKeyResult: result2 = evaluateIndexOrKey(key)
            If mItemCount > 0 And (result2.Index >= 0 Or result2.Index <= (mItemCount - 1)) Then
                Err.Raise 457, "DictCollection", "This key is already associated with an element of this collection"
            Else
                Err.Raise 13, "DictCollection", "Type mismatch": Exit Function
            End If
        Else
            If (TypeName(key) <> "String") And (Not IsMissing(key)) Then Err.Raise 13, "DictCollection", "Type mismatch"
        End If
    End If
    If mEmulateScriptingDictionary Then
        result.key = convertScriptingDictionaryToDictCollectionKey(key)
        result.WasInvalid = False: result.WasString = True ' explicitly set to string
    ElseIf mEmulateCollection Then
        result = evaluateKey(convertCollectionToDictCollectionKey(key))
    Else
        If IsMissing(key) Then key = ""
        result = evaluateKey(key)
    End If
    If result.WasInvalid Then
        If mThrowErrors Then Err.Raise 13, "DictCollection", "Type mismatch"
    End If
    internalAddItem Item, result.key, AtIndex
    If mEmulateScriptingDictionary Then storeScriptingDictionaryKeyInformation key, result.key
ExitFunction:
    Set Add2 = Me
End Function

' Adds item with given key ("" = no key) at given index (if an item exists at index, it will be pushed one up)
' This function is emulation independent (with the exception of throwing compatible errors)
Private Sub internalAddItem(Optional Item As Variant = Empty, Optional key As String = "", Optional AtIndex As Long = -1)
    Dim NewKey As FindKeyResult, newItemIndex As Long, lastItemIndex As Long, newUpperBound As Long, keyPassed As Boolean, indexPassed As Boolean
    keyPassed = Len(key) > 0 Or mEmulateScriptingDictionary
    If keyPassed Then
        If mKeyCount = -1 Then initializeKeyArray 'convert array to collection if key is being used
        If mLazySorting And Not mThrowErrors Then
            If AtIndex > -1 Then
                ' key was passed and item will be put somewhere in item array -> sort before because key collisions can't be fixed after this insert
                If Not mKeysAreSortedAndUnique And mKeyCount > 1 Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
                NewKey = internalFindKeyIndex(key, mCompareMode) 'find key
            Else
                ' put key to position after last
                NewKey.Index = mKeyCount
                NewKey.Exact = False
                mKeysAreSortedAndUnique = False
            End If
        ElseIf mLazySorting And mThrowErrors Then
            ' key was passed and we need to check if key already exists -> sort before because otherwise we can't throw an error
            If Not mKeysAreSortedAndUnique And mKeyCount > 1 Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
            NewKey = internalFindKeyIndex(key, mCompareMode) 'find key
        Else
            NewKey = internalFindKeyIndex(key, mCompareMode) 'find key
        End If
    Else
        NewKey.Exact = False
    End If
    If NewKey.Exact Then
        ' we only get here if key was passed and key already exists
        If mThrowErrors Then
            'throw error because of Scripting.Dictionary emulation
            Err.Raise 457, "DictCollection", "This key is already associated with an element of this collection"
            Exit Sub
        Else
            Dim existingItemIndex As Long: existingItemIndex = allKeyItemIndexes(NewKey.Index)
            If AtIndex = -1 Or (AtIndex = existingItemIndex) Then
                ' key exists, item index is the same, replace item
                If IsObject(Item) Then Set allItems(existingItemIndex) = Item Else allItems(existingItemIndex) = Item
            Else
                Dim moveItemsUpwards As Boolean, upper As Long, lower As Long
                If AtIndex > (mItemCount - 1) Then increaseItemArrayTo AtIndex
                moveItemsUpwards = (existingItemIndex > AtIndex)
                If moveItemsUpwards Then ' copy items up
                    lower = AtIndex: upper = existingItemIndex - 1 ' existing item will be overwritten by the one before it
                    copyItemsUp lower, upper, 1, True
                Else ' copy items down
                    lower = existingItemIndex + 1: upper = AtIndex ' existing item will be overwritten by the one after it
                    copyItemsDown lower, upper, 1, True
                End If
                If IsObject(Item) Then Set allItems(AtIndex) = Item Else allItems(AtIndex) = Item
                allItemKeyIndexes(AtIndex) = NewKey.Index ' set item-to-key pointer
                allKeyItemIndexes(NewKey.Index) = AtIndex ' set key-to-item-pointer
            End If
            allKeys(NewKey.Index) = key ' overwrite key just in case it is different than old key
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
                Dim dc2 As DictCollection: Set dc2 = Item
                CopySubCollectionChainingSettings Me, dc2
            End If
        Else
            allItems(newItemIndex) = Item
        End If
        If keyPassed Then
            allItemKeyIndexes(newItemIndex) = NewKey.Index
            ' STEP 2: insert key in keyArray
            insertKey key, NewKey.Index, newItemIndex
            ' invalidate sorting because 1) new key might be at wrong index and 2) new key might be a duplicate of previously added key
            If mLazySorting Then mKeysAreSortedAndUnique = False
        Else
            allItemKeyIndexes(newItemIndex) = -1 ' set to -1 = no key exists for this item
        End If
    End If
End Sub

' inserts key at given index after copying all existing keys >= KeyIndex one up
Private Function insertKey(key As String, KeyIndex As Long, ItemIndex As Long)
    increaseKeyArray
    If KeyIndex < (mKeyCount - 1) Then copyKeysUp KeyIndex, mKeyCount - 2, 1, True
    allKeys(KeyIndex) = key
    allKeyItemIndexes(KeyIndex) = ItemIndex
End Function

' Removes an item at a given index or key; chainable
Public Function Remove(IndexOrKey As Variant) As DictCollection
    If mItemCount < 0 Then GoTo ExitFunction
    Dim result As EvaluateIndexOrKeyResult, found As Boolean, key2 As String
    If mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
    If mEmulateScriptingDictionary Then
        key2 = convertScriptingDictionaryToDictCollectionKey(IndexOrKey)
        result = evaluateIndexOrKey(key2)
    ElseIf mEmulateCollection Then
        If TypeName(IndexOrKey) = "String" Then
            result = evaluateIndexOrKey(convertCollectionToDictCollectionKey(IndexOrKey))
        Else
            result = evaluateIndexOrKey(IndexOrKey)
        End If
    Else
        result = evaluateIndexOrKey(IndexOrKey)
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
    Dim KeyIndex As Long, hasKeys As Boolean
    hasKeys = mKeyCount > 0
    KeyIndex = allItemKeyIndexes(result.Index)
    If (result.Index = (mItemCount - 1)) Then
        'last item, do nothing
    Else
        ' copy all items after found index one down
        copyItemsDown result.Index + 1, mItemCount - 1, 1, hasKeys
    End If
    decreaseItemArray
    If hasKeys And KeyIndex > -1 Then
        If KeyIndex = (mKeyCount - 1) Then
            ' last item, do nothing
        Else
            ' copy keys one down
            copyKeysDown KeyIndex + 1, (mKeyCount - 1), 1, True
        End If
        decreaseKeyArray
        If mEmulateScriptingDictionary Then
            Dim objectKeyIndex As Long: objectKeyIndex = getScriptingDictionaryObjectKeyIndex(key2)
            If objectKeyIndex > -1 Then removeScriptingDictionaryObjectKey objectKeyIndex
            If isScriptingDictionaryStoredKeyType(key2) Then mScriptingDictionaryStoredKeyTypes.Remove key2
        End If
    End If
ExitFunction:
    Set Remove = Me
End Function

' Removes all existing items and resets the DictCollection to array; chainable
Public Function RemoveAll() As DictCollection
    initializeItemArray
    uninitializeKeyArray
    If mEmulateScriptingDictionary Then
        initializeDictionaryStoredKeyTypes
        initializeDictionaryObjectKeyArray
    Else
        uninitializeDictionaryStoredKeyTypes
        uninitializeDictionaryObjectKeyArray
    End If
    Set RemoveAll = Me
End Function

' Returns true if key exists, false of not
' variant key is compatible with Scripting.Dictionary
Public Function Exists(key As Variant) As Boolean
    Dim retVal As Boolean, result As EvaluateKeyResult
    retVal = False
    If mKeyCount > -1 Then
        If mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
        If mEmulateScriptingDictionary Then
            result = evaluateKey(convertScriptingDictionaryToDictCollectionKey(key))
        ElseIf mEmulateCollection Then
            result = evaluateKey(convertCollectionToDictCollectionKey(key))
        Else
            result = evaluateKey(key)
        End If
        If Not result.WasInvalid Then retVal = internalFindKeyIndex(result.key, mCompareMode).Exact
    End If
    Exists = retVal
End Function
' ======================== END: ITEM, ITEMAT, ITEMOF, ADD, INSERT, ADD2, EXISTS, REMOVE, REMOVEALL ==================================================


' ======================== START: COUNT, KEYCOUNT, INDEXOFKEY, INDEXOFKEY2, KEYOFITEMAT, KEY, ITEMHASKEY ============================================
' returns the number of stored items (including Empty ones)
Public Property Get Count() As Long
    ' sort keys and remove duplicate keys + their items, then count afterwards
    If mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
    Count = IIf(mItemCount = -1, 0, mItemCount)
End Property

' returns the number of stored keys
Public Property Get KeyCount() As Long
    ' sort keys and remove duplicate keys + their items, then count afterwards
    If mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
    KeyCount = IIf(mKeyCount = -1, 0, mKeyCount)
End Property

' Returns index of a given key. If ThrowErrors=False, returns -1 if key was not found
' Keys that can be converted to String will be converted and matched against stored keys
Public Property Get IndexOfKey(key As Variant) As Long
    Dim retVal As Long: retVal = -1
    If mItemCount < 1 Then
        ' Emulate Collection behavior
        If mThrowErrors And Not mEmulateScriptingDictionary Then Err.Raise 5, "DictCollection", "Invalid procedure call or argument"
        GoTo ExitFunction
    End If
    Dim result As EvaluateKeyResult
    If mEmulateScriptingDictionary Then
        result.key = convertScriptingDictionaryToDictCollectionKey(key)
    ElseIf mEmulateCollection Then
        ' Emulate Collection behavior of accepting only String
        If TypeName(key) <> "String" Then
            If mThrowErrors Then Err.Raise 13, "DictCollection", "Keys of type " & TypeName(key) & " are not supported"
            GoTo ExitFunction
        End If
        result.key = convertCollectionToDictCollectionKey(key)
    Else
        result = evaluateKey(key)
        If result.WasInvalid Then
            If mThrowErrors Then Err.Raise 13, "DictCollection", "Keys of type " & TypeName(key) & " are not supported"
            GoTo ExitFunction
        End If
        If mThrowErrors And (IsEmpty(key) Or Len(result.key) = 0) Then Err.Raise 5, "DictCollection", "Empty or zero length keys are not supported": GoTo ExitFunction
    End If
    If mKeyCount < 1 Then
        ' Emulate Collection behavior after not throwing type errors
        If mThrowErrors And Not mEmulateScriptingDictionary Then Err.Raise 5, "DictCollection", "Invalid procedure call or argument"
        GoTo ExitFunction
    End If
    If mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
    Dim fkr As FindKeyResult
    fkr = internalFindKeyIndex(result.key, mCompareMode)
    If fkr.Exact Then
        Dim ItemIndex As Long
        ItemIndex = allKeyItemIndexes(fkr.Index)
        If ItemIndex < 0 Or ItemIndex > (mItemCount - 1) Then
            Debug.Print "!!!!!! Bug found in DictCollection function IndexOfKey(" & key & "): key points to index=" & ItemIndex & " outside of item array !!!!!!!"
            If mThrowErrors Then Err.Raise 9, "DictCollection", "Subscript out of range": Exit Property
        End If
        retVal = ItemIndex
    Else
        If mThrowErrors Then Err.Raise 5, "DictCollection", "Key does not exist"
        GoTo ExitFunction
    End If
ExitFunction:
    If retVal = -1 Then IndexOfKey = -1 Else IndexOfKey = retVal + mExternalIndexOffset
End Property

' UNTESTED
' Alternative (and fast) function for looking up a key's index without any type conversion
' Returns -1 if key is not found and .ThrowErrors=False
Public Property Get IndexOfKey2(key As String) As Long
    Dim retVal As Long: retVal = -1
    If mKeyCount < 1 Or mItemCount < 1 Then Exit Property
    Dim foundKey As FindKeyResult
    ' sort keys if needed
    If mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
    If mEmulateScriptingDictionary Then
        foundKey = internalFindKeyIndex(convertScriptingDictionaryToDictCollectionKey(key), mCompareMode)
    ElseIf mEmulateCollection Then
        foundKey = internalFindKeyIndex(convertCollectionToDictCollectionKey(key), mCompareMode)
    Else
        foundKey = internalFindKeyIndex(key, mCompareMode)
    End If
    If foundKey.Exact Then
        retVal = allKeyItemIndexes(foundKey.Index)
    Else
        If mThrowErrors Then Err.Raise 5, "DictCollection", "Key does not exist"
    End If
    IndexOfKey2 = retVal
End Property

' Returns key of given item index. If ThrowErrors=False, returns empty string if key was not found
Public Property Get KeyOfItemAt(Index As Long) As Variant
    KeyOfItemAt = Empty ' default value if not found
    If mKeyCount < 1 Or mItemCount < 1 Then
        ' Emulate Collection behavior
        If mThrowErrors Then Err.Raise 5, "DictCollection", "Invalid procedure call or argument"
        Exit Property
    End If
    If mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
    Dim internalIndex As Long: internalIndex = Index - mExternalIndexOffset
    If internalIndex < 0 Or internalIndex >= mItemCount Then
        If mThrowErrors Then Err.Raise 9, "DictCollection", "No item at index " & Index
        Exit Property
    End If
    Dim KeyIndex As Long
    KeyIndex = allItemKeyIndexes(internalIndex)
    If KeyIndex < 0 Or KeyIndex >= mKeyCount Then Exit Property
    Dim key As String: key = allKeys(KeyIndex)
    If mEmulateScriptingDictionary Then
        If isStoredKeyOfTypeObject(key) Then Set KeyOfItemAt = convertDictCollectionToScriptingDictionaryKey(key) Else KeyOfItemAt = convertDictCollectionToScriptingDictionaryKey(key)
    ElseIf mEmulateCollection Then
        KeyOfItemAt = convertDictCollectionToCollectionKey(key)
    Else
        KeyOfItemAt = key
    End If
End Property

' Sets the key of an item at a given index
' If the index is greater than the current item storage size and .ThowError=False, the item array will be expanded to the given index
' If the new key value is already used, the key of the item that was previously associated with that key be set to none but the item will not be removed (unlike .Key)
Public Property Let KeyOfItemAt(Index As Long, Value As Variant)
    Dim index2 As Long
    index2 = Index - mExternalIndexOffset
    If index2 < 0 Then
        If mThrowErrors Then Err.Raise 9, "DictCollection", "Index must be greater than or equal " & mExternalIndexOffset
        Exit Property
    End If
    If index2 >= mItemCount Then
        If mThrowErrors Then
            Err.Raise 9, "DictCollection", "No item at index " & index2
        Else
            ' increase item array to given index
            increaseItemArrayTo index2
        End If
    End If
    internalSetKey index2, Value, False
End Property

' returns the key as String, when emulating Dictionary/Collection, returns internal representation of the key
' Index is always zero-based
Public Function KeyOfItemAtAsString(Index As Long) As String
    Dim retVal As String: retVal = ""
    If mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
    If Index > -1 And Index <= (mItemCount - 1) Then
        If allItemKeyIndexes(Index) > -1 Then retVal = allKeys(allItemKeyIndexes(Index))
    Else
        If mThrowErrors Then Err.Raise 9, "DictCollection", "No item found at index " & Index
    End If
    KeyOfItemAtAsString = retVal
End Function

' Changes an existing key to another value
' If the new key already exists and ThrowErrors=False, the item previously associated with that key will be removed (unlike .KeyOfItemAt)
' If the new key value can be converted to String, it will be set to the value converted to String
' If the new key value is empty or "" then the existing key will be removed or set to none
Public Property Let key(Previous As Variant, Value As Variant)
    Dim oldKeyEvalResult As EvaluateKeyResult, oldKeyFindResult As FindKeyResult, oldItemIndex As Long, newKeyEvalResult As EvaluateKeyResult
    ' convert old key to internal key and check if new key and old key are the same, throw errors if necessary
    If mEmulateScriptingDictionary Then
        If IsObject(Previous) And IsObject(Value) Then
            If Previous Is Value Then Exit Property ' two objects are equal, exit
        ElseIf IsNumeric(Previous) And IsNumeric(Value) And VarType(Previous) <> VarType(Value) Then
            If Previous = Value Then Exit Property ' two numbers are equal, exit
        ElseIf IsDate(Previous) And IsDate(Value) Then
            If Previous = Value Then Exit Property ' two numbers are equal, exit
        ElseIf (IsNumeric(Previous) And VarType(Value) = VbVarType.vbDate) Or (VarType(Previous) = VbVarType.vbDate And IsNumeric(Value)) Then
            ' since date is not numeric, we have to check this differently
            If CDbl(Previous) = CDbl(Value) Then Exit Property ' number and date are equal, exit
        End If
        oldKeyEvalResult.key = convertScriptingDictionaryToDictCollectionKey(Previous)
        oldKeyEvalResult.WasInvalid = False: oldKeyEvalResult.WasString = True ' explicitly set to string
    ElseIf mEmulateCollection Then
        oldKeyEvalResult = evaluateKey(convertCollectionToDictCollectionKey(Previous))
    Else
        oldKeyEvalResult = evaluateKey(Previous)
    End If
    
    If oldKeyEvalResult.WasInvalid Or (mEmulateCollection And (Not oldKeyEvalResult.WasString Or IsMissing(Previous))) Then
        If mThrowErrors Then Err.Raise 13, "DictCollection", "Keys of type " & TypeName(Value) & " are not supported"
        Exit Property
    ElseIf oldKeyEvalResult.key = "" Then
        ' since we don't support empty keys, emulate Scripting.Dictionary behavior for nonexisting keys
        If mThrowErrors Then Err.Raise 32811, "DictCollection", "Method 'Key' of 'DictCollection' failed."
        Exit Property
    End If
        
    If mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
    oldKeyFindResult = internalFindKeyIndex(oldKeyEvalResult.key, mCompareMode)
    
    If Not oldKeyFindResult.Exact Then
        ' previous key not found, emulate Scripting.Dictionary behavior for nonexisting keys
        If mThrowErrors Then Err.Raise 32811, "DictCollection", "Method 'Key' of 'DictCollection' failed."
        Exit Property
    End If
    ' if we got here, we know that (1) previous key exists and (2) key / item arrays are not empty and (3) the new key is different from the old one
    oldItemIndex = allKeyItemIndexes(oldKeyFindResult.Index)
    ' set key using internal function
    internalSetKey oldItemIndex, Value, True
End Property

' sets a new key for a given item index; if the new key already exists and ThrowErrors=False, its item will be removed
' If the new key value can be converted to String, it will be set to the value converted to String
' If the new key value is empty or "" then the existing key will be removed or set to none
Private Sub internalSetKey(ItemIndex As Long, NewKey As Variant, RemoveItemPreviouslyAssociatedWithNewKey As Boolean)
    Dim newKeyEvalResult As EvaluateKeyResult, oldKeyIndex As Long, oldKey2 As String
    
    oldKeyIndex = allItemKeyIndexes(ItemIndex)
    If oldKeyIndex > -1 Then oldKey2 = allKeys(oldKeyIndex)

    If mEmulateScriptingDictionary Then
        newKeyEvalResult.key = convertScriptingDictionaryToDictCollectionKey(NewKey)
        newKeyEvalResult.WasInvalid = False: newKeyEvalResult.WasString = True ' explicitly set to string
    ElseIf mEmulateCollection Then
        newKeyEvalResult = evaluateKey(convertCollectionToDictCollectionKey(NewKey))
    Else
        newKeyEvalResult = evaluateKey(NewKey)
    End If
    If newKeyEvalResult.WasInvalid Or (mEmulateCollection And (Not newKeyEvalResult.WasString Or IsMissing(NewKey))) Then
        If mThrowErrors Then Err.Raise 13, "DictCollection", "Keys of type " & TypeName(NewKey) & " are not supported"
        Exit Sub
    End If
    If (mItemCount < 1) Then
        If mThrowErrors And Not mEmulateScriptingDictionary Then Err.Raise 5, "DictCollection", "Invalid procedure call or argument"
        Exit Sub
    End If
    
    If newKeyEvalResult.key = "" And oldKeyIndex = -1 Then
        ' do nothing, old and new keys are not set and thus the same
    ElseIf newKeyEvalResult.key = "" And oldKeyIndex > -1 Then
        ' new key is empty -> remove found previous (old) key
        If oldKeyIndex = (mItemCount - 1) Then
            ' last item, do nothing
        Else
            copyKeysDown oldKeyIndex + 1, (mKeyCount - 1), 1, True ' copy keys one down
        End If
        decreaseKeyArray
        ' and set keyIndex for corresponding item to -1 (= no key)
        allItemKeyIndexes(ItemIndex) = -1
    Else
        Dim newKeyFindResult As FindKeyResult
        ' check if new key is already used before replacing found previous (old) key
        newKeyFindResult = internalFindKeyIndex(newKeyEvalResult.key, mCompareMode)
        If newKeyFindResult.Exact Then
            ' new key exists!
            If newKeyFindResult.Index = oldKeyIndex Then
                ' new key and old key are the same (when for instance .CompareMode = 1 -> case insensitive comparison), set key to new one
                allKeys(oldKeyIndex) = newKeyEvalResult.key
            Else
                ' Emulate Scripting.Dictionary and Collection behavior
                If mThrowErrors Then Err.Raise 457, "DictCollection", "This key is already associated with an element of this collection": Exit Sub
                ' remove the item previously associates with the new key and set the new key for the item previously associated with the old key
                Dim itemIndexThatPreviouslyHadNewKey As Long: itemIndexThatPreviouslyHadNewKey = allKeyItemIndexes(newKeyFindResult.Index)
                If oldKeyIndex > -1 Then
                    ' item had a key previously
                    ' (1) remove old key and adjust newKeyFindResult.index if necessary (if key was moved down)
                    copyKeysDown oldKeyIndex + 1, (mKeyCount - 1), 1, True ' copy keys one down
                    If newKeyFindResult.Index > oldKeyIndex Then newKeyFindResult.Index = newKeyFindResult.Index - 1
                    decreaseKeyArray
                End If
                ' (2) write new item index to key-to-item pointer
                allKeyItemIndexes(newKeyFindResult.Index) = ItemIndex
                ' (3) write the new key just in case it is not exactly the same as the found one (case insensitive matching)
                allKeys(newKeyFindResult.Index) = newKeyEvalResult.key
                If mEmulateScriptingDictionary Then
                    ' also update the key in the type storage
                    If isScriptingDictionaryStoredKeyType(oldKey2) Then mScriptingDictionaryStoredKeyTypes.key(oldKey2) = newKeyEvalResult.key
                End If
                ' (4) replace key of item associated with old key with new key
                allItemKeyIndexes(ItemIndex) = newKeyFindResult.Index
                If RemoveItemPreviouslyAssociatedWithNewKey Then
                    ' (5A) if called from .Key(), RemoveItemPreviouslyAssociatedWithNewKey=True and we can remove the item previously associated with the new key
                    copyItemsDown itemIndexThatPreviouslyHadNewKey + 1, (mItemCount - 1), 1, True
                    decreaseItemArray
                Else
                    ' (5B) if called by .KeyOfItemAt(), RemoveItemPreviouslyAssociatedWithNewKey=False and we have to
                    ' set the key of item previously associated with the new key to none (its item-to-key pointer to -1)
                    allItemKeyIndexes(itemIndexThatPreviouslyHadNewKey) = -1
                End If
            End If
        Else
            ' new key not found -> newKeyFindResult.index is the index where the new key should be inserted
                
            If oldKeyIndex > -1 And ((oldKeyIndex = newKeyFindResult.Index) Or (oldKeyIndex = newKeyFindResult.Index - 1)) Then
                ' new key should be inserted at the same index as old one -> new key overwrites old one
                allKeys(oldKeyIndex) = newKeyEvalResult.key
            ElseIf oldKeyIndex > newKeyFindResult.Index Then
                ' new key should be inserted before old key -> copy all keys from new key index to old key index -1 one up
                copyKeysUp newKeyFindResult.Index, oldKeyIndex - 1, 1, True
                allKeys(newKeyFindResult.Index) = newKeyEvalResult.key
                allKeyItemIndexes(newKeyFindResult.Index) = ItemIndex
                allItemKeyIndexes(ItemIndex) = newKeyFindResult.Index
            ElseIf oldKeyIndex < newKeyFindResult.Index Then
                If oldKeyIndex = -1 Then
                    If mKeyCount < 0 Then initializeKeyArray
                    ' old key was set to none, newKeyFindResult.Index is always at least 0
                    increaseKeyArray
                    ' insert new key at found index:
                    ' 1) copy all previous keys from found index upwards one up
                    copyKeysUp newKeyFindResult.Index, mKeyCount - 2, 1, True
                    ' 2) insert the new key and set all pointers
                    allKeys(newKeyFindResult.Index) = newKeyEvalResult.key
                    allKeyItemIndexes(newKeyFindResult.Index) = ItemIndex
                    allItemKeyIndexes(ItemIndex) = newKeyFindResult.Index
                Else
                    ' new key should be inserted after old key -> copy all keys between old and new key one down
                    copyKeysDown oldKeyIndex + 1, newKeyFindResult.Index - 1, 1, True
                    ' new key's index will be newKeyFindResult.index - 1 since old key is now gone
                    allKeys(newKeyFindResult.Index - 1) = newKeyEvalResult.key
                    allKeyItemIndexes(newKeyFindResult.Index - 1) = ItemIndex
                    allItemKeyIndexes(ItemIndex) = newKeyFindResult.Index - 1
                End If
            End If
            If mEmulateScriptingDictionary Then
                ' check if old key was object and if yes, remove it
                Dim oldObjectKeyIndex As Long, objectKeyIndex As Long: oldObjectKeyIndex = getScriptingDictionaryObjectKeyIndex(oldKey2)
                If oldObjectKeyIndex > -1 Then removeScriptingDictionaryObjectKey oldObjectKeyIndex
                ' check if new key is object and if yes, add it
                Dim newObjectKeyIndex As Long: newObjectKeyIndex = getScriptingDictionaryObjectKeyIndex(newKeyEvalResult.key)
                If newObjectKeyIndex > -1 Then UtilAddArrayValue mScriptingDictionaryObjectKeys, NewKey
                ' check if old key is stored key type and if yes, remove it
                If isScriptingDictionaryStoredKeyType(oldKey2) Then mScriptingDictionaryStoredKeyTypes.Remove (oldKey2)
                ' check if new key is number type and if yes, remember it
                If isScriptingDictionaryStoredKeyType(newKeyEvalResult.key) Then mScriptingDictionaryStoredKeyTypes(newKeyEvalResult.key) = TypeName(NewKey)
            End If
        End If
    End If
End Sub

' returns true if item at given Index has a key and false if not
Public Function ItemHasKey(ItemIndex As Long) As Boolean
    Dim retVal As Boolean, index2 As Long: retVal = False
    index2 = ItemIndex - mExternalIndexOffset
    If index2 > -1 And index2 <= (mItemCount - 1) Then
        retVal = (allItemKeyIndexes(index2) > -1)
    Else
        If mThrowErrors Then Err.Raise 9, "DictCollection", "No item found at index " & ItemIndex
    End If
    ItemHasKey = retVal
End Function
' ======================== END: COUNT, KEYCOUNT, INDEXOFKEY, INDEXOFKEY2, KEYOFITEMAT, KEY, ITEMHASKEY ==============================================


' ======================== START: KEYS AND ITEMS, FINDKEY FUNCTIONS, TOARRAY ========================================================================
' Returns all keys as variant array in the order of the items (as they were added). The value Empty means this item has no associated key.
' Example1: for items [[item="A",key="key1"],[item="B"]] the keys would be [["key1"],[Empty]]
' If IncludeEmptyKeys = False, Empty values are omitted and the array contains all stored keys in the order of the items.
' Variant array is compatible with Scripting.Dictionary
Public Property Get Keys(Optional IncludeEmptyKeys As Boolean = True, Optional IncludeKeysWithEmptyItems As Boolean = True) As Variant
    Keys = internalGetKeys(IncludeEmptyKeys, IncludeKeysWithEmptyItems, False, False)
End Property

' Returns all keys in sorted order as two-dimensional Variant array with array(ItemIndex,0) = Item and array(ItemIndex,1) = Key
' Variant array is compatible with Scripting.Dictionary emulation
Public Property Get SortedKeys(Optional IncludeEmptyItems As Boolean = True) As Variant
    SortedKeys = internalGetKeys(False, IncludeEmptyItems, True, False)
End Property

' UNTESTED
' Returns all sorted keys and their items as two-dimensional Variant array with array(ItemIndex,0) = Key and array(ItemIndex,1) = Item
' Returns a one-dimensional empty array if DictCollection contains no keys
Public Property Get SortedKeysAndItems(Optional IncludeEmptyKeys As Boolean = True, Optional IncludeEmptyItems As Boolean = True) As Variant
    SortedItemsAndKeys = internalGetKeys(IncludeEmptyKeys, IncludeEmptyItems, True, True)
End Property

' UNTESTED
' Returns all items (in the order they where added) and keys as two-dimensional Variant array with array(ItemIndex,0) = Key and array(ItemIndex,1) = Item
Public Function ToArray() As Variant
    ToArray = internalGetKeys(True, True, False, True)
End Function

' UNTESTED
' Returns two-dimensional array containing all items that are DictCollections (= subcollections) as rows and their items as columns.
' ColHeaderKeys define the keys of the subcollections that should be returned and their order.
Public Function ToColumns(ColHeaderKeys As Variant, Optional FlattenSubCollections As Boolean = False) As Variant
    Dim retVal As Variant, subDc As DictCollection, subSubDc As DictCollection, colCount As Long, RowCount As Long, i As Long, j As Long
    Dim chLower As Long, chUpper As Long, key As String, val As Variant, val2 As String
    RowCount = mItemCount: chLower = LBound(ColHeaderKeys): chUpper = UBound(ColHeaderKeys)
    colCount = chUpper - chLower + 1
    If mItemCount = 0 Then
        retVal = Empty
    Else
        ReDim retVal(RowCount - 1, colCount - 1)
        For i = 0 To mItemCount - 1
            ' collect values only for items that are DictCollections
            If TypeName(allItems(i)) = "DictCollection" Then
                Set subDc = allItems(i)
                If FlattenSubCollections Then Set subDc = subDc.FlattenToNew
                For j = 0 To colCount - 1
                    key = CStr(ColHeaderKeys(chLower + j))
                    ' if key exists, store value in retVal
                    If subDc.IndexOfKey2(key) >= 0 Then
                        val = subDc.ItemOf(key)
                        If TypeName(val) = "DictCollection" Then
                            ' we never get here with FlattenSubCollections=True
                            Set subSubDc = val
                            Set subSubDc = subSubDc.FlattenToNew
                            val = subSubDc.Items
                            val = Join(val, ",")
                        End If
                        retVal(i, j) = val
                    End If
                Next
            Else
                ' not a DictCollection, ignore
            End If
        Next
    End If
    ToColumns = retVal
End Function

' UNTESTED
' Returns two-dimensional array containing all items that are DictCollections (= subcollections) as rows and their items as columns.
' The first row is a list of all keys (columns) found in subcollections. All unkeyed values in subcollections will be omitted.
' If a subcollection does not have a value for a key (=column), the value will be set to FillValue.
' The first column is a list of the keys of the subcollections if the subcollections themselves are stored with keys.
Public Function ToTable(FillValue As Variant, Optional FlattenSubCollections As Boolean = False) As Variant
    Dim retVal As Variant, subDc As DictCollection, subSubDc As DictCollection, colCount As Long, colOffset As Long, RowCount As Long, i As Long, j As Long
    Dim key As String, val As Variant, val2 As String, keyDc As New DictCollection, flattenedDcs As New DictCollection
    If mItemCount = 0 Then
        retVal = Empty
    Else
        ' count rows and collect all subcollection keys into key collection
        RowCount = 0
        For i = 0 To mItemCount
            If TypeName(allItems(i)) = "DictCollection" Then
                Set subDc = allItems(i)
                If FlattenSubCollections Then
                    Set subDc = subDc.FlattenToNew
                    ' store flattened dc in cache
                    flattenedDcs.Insert subDc, i
                End If
                For j = 0 To subDc.KeyCount: keyDc.Add subDc.KeyAtKeyIndex(i), Empty: Next
                RowCount = RowCount + 1
            End If
        Next
        colCount = keyDc.KeyCount
        If mKeyCount > 0 Then
            colCount = colCount + 1 ' add frist col for keys
            colOffset = 1 ' values start at column index 1
        Else
            colOffset = 0 ' values start at column index 0
        End If
        ReDim retVal(RowCount - 1, colCount - 1)
        For i = 0 To mItemCount - 1
            ' if there are keys, copy key to firct column or set it to empty if item has no key
            If mKeyCount > 0 And allItemKeyIndexes(i) > -1 Then retVal(i, 0) = allKeys(allItemKeyIndexes(i)) Else retVal(i, 0) = Empty
            ' collect values only for items that are DictCollections
            If TypeName(allItems(i)) = "DictCollection" Then
                ' set subDc or retrieve flattened dc from cache
                If FlattenSubCollections Then Set subDc = flattenedDcs.ItemAt(i) Else Set subDc = allItems(i)
                For j = 0 To keyDc
                    key = subDc.KeyOfItemAt(j)
                    If subDc.IndexOfKey2(key) >= 0 Then
                        ' it has a value
                        val = subDc.ItemOf(key)
                        If TypeName(val) = "DictCollection" Then
                            ' we never get here with FlattenSubCollections=True
                            Set subSubDc = val
                            Set subSubDc = subSubDc.FlattenToNew
                            val = subSubDc.Items
                            val = Join(val, ",")
                        End If
                    Else
                        val = FillValue
                    End If
                    retVal(i, j + colOffset) = val
                Next
            Else
                ' not a DictCollection, igore
            End If
        Next
    End If
    ToTable = retVal
End Function

' internal function that builds one/two-dimensional output array
Private Function internalGetKeys(Optional IncludeEmptyKeys As Boolean = True, Optional IncludeEmptyItems As Boolean = True, Optional SortedByKeys As Boolean = False, Optional IncludeItems As Boolean = False) As Variant
    Dim retVal() As Variant, retValIndex As Long, retValCount As Long, ItemIndex As Long, KeyIndex As Long, added As Boolean
    If mKeyCount < 1 Then retVal = Array(): GoTo ExitFunction
    
    ' sort keys if needed
    If mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
    
    ' count number of output values
    retValCount = 0
    If IncludeEmptyItems Then
        If IncludeEmptyKeys Then retValCount = mItemCount Else retValCount = mKeyCount
    Else
        ' no empty items
        If IncludeEmptyKeys Then
            ' include nonexisting keys, run over items
            For ItemIndex = 0 To (mItemCount - 1)
                If IsObject(allItems(ItemIndex)) Then
                    retValCount = retValCount + 1
                ElseIf Not IsEmpty(allItems(ItemIndex)) Then
                    retValCount = retValCount + 1
                End If
            Next
        Else
            ' only existing keys, run over keys
            For KeyIndex = 0 To (mKeyCount - 1)
                ItemIndex = allKeyItemIndexes(KeyIndex)
                If IsObject(allItems(ItemIndex)) Then
                    retValCount = retValCount + 1
                ElseIf Not IsEmpty(allItems(ItemIndex)) Then
                    retValCount = retValCount + 1
                End If
            Next
        End If
    End If
    
    retValIndex = 0
    If IncludeItems Then
        ' two-dimensional array with keys and items
        ReDim retVal(0 To (retValCount - 1), 1)
        If SortedByKeys Then
            ' include items, sorted by keys, run over all keys and add all items/keys that match criteria
            For KeyIndex = 0 To (mKeyCount - 1)
                added = False
                ItemIndex = allKeyItemIndexes(KeyIndex)
                If IsObject(allItems(ItemIndex)) Then
                    Set retVal(retValIndex, 1) = allItems(ItemIndex): added = True
                ElseIf IncludeEmptyItems Or (Not IsEmpty(allItems(ItemIndex)) And Not IncludeEmptyItems) Then
                    retVal(retValIndex, 1) = allItems(ItemIndex): added = True
                End If
                If added Then
                    ' set key
                    If mEmulateScriptingDictionary Then
                        If isStoredKeyOfTypeObject(allKeys(KeyIndex)) Then
                            Set retVal(retValIndex, 0) = convertDictCollectionToScriptingDictionaryKey(allKeys(KeyIndex))
                        Else
                            retVal(retValIndex, 0) = convertDictCollectionToScriptingDictionaryKey(allKeys(KeyIndex))
                        End If
                    ElseIf mEmulateCollection Then
                        retVal(retValIndex, 0) = convertDictCollectionToCollectionKey(allKeys(KeyIndex))
                    Else
                        retVal(retValIndex, 0) = allKeys(KeyIndex)
                    End If
                    retValIndex = retValIndex + 1
                End If
            Next
            ' add items without keys at end
            If IncludeEmptyKeys Then
                ' add items without keys
                For ItemIndex = 0 To (mItemCount - 1)
                    added = False
                    If allItemKeyIndexes(ItemIndex) = -1 Then
                        If IsObject(allItems(ItemIndex)) Then
                            Set retVal(retValIndex, 1) = allItems(ItemIndex): added = True
                        ElseIf IncludeEmptyItems Or (Not IsEmpty(allItems(ItemIndex)) And Not IncludeEmptyItems) Then
                            retVal(retValIndex, 1) = allItems(ItemIndex): added = True
                        End If
                    End If
                    If added Then
                        retVal(retValIndex, 0) = Empty
                        retValIndex = retValIndex + 1
                    End If
                Next
            End If
        Else
            ' include items, no sorting, run over items and add items/keys that match criteria
            For ItemIndex = 0 To (mItemCount - 1)
                added = False
                KeyIndex = allItemKeyIndexes(ItemIndex)
                If IncludeEmptyKeys Or (Not IncludeEmptyKeys And KeyIndex > -1) Then
                    If KeyIndex > -1 Then
                        ' key exists
                        If IsObject(allItems(ItemIndex)) Then
                            Set retVal(retValIndex, 1) = allItems(ItemIndex): added = True
                        ElseIf IncludeEmptyItems Or (Not IsEmpty(allItems(ItemIndex)) And Not IncludeEmptyItems) Then
                            retVal(retValIndex, 1) = allItems(ItemIndex): added = True
                        End If
                        If added Then
                            ' set key
                            If mEmulateScriptingDictionary Then
                                If isStoredKeyOfTypeObject(allKeys(KeyIndex)) Then
                                    Set retVal(retValIndex, 0) = convertDictCollectionToScriptingDictionaryKey(allKeys(KeyIndex))
                                Else
                                    retVal(retValIndex, 0) = convertDictCollectionToScriptingDictionaryKey(allKeys(KeyIndex))
                                End If
                            ElseIf mEmulateCollection Then
                                retVal(retValIndex, 0) = convertDictCollectionToCollectionKey(allKeys(KeyIndex))
                            Else
                                retVal(retValIndex, 0) = allKeys(KeyIndex)
                            End If
                        End If
                    Else
                        ' key does not exist
                        If IsObject(allItems(ItemIndex)) Then
                            Set retVal(retValIndex, 1) = allItems(ItemIndex): added = True
                        ElseIf IncludeEmptyItems Or (Not IsEmpty(allItems(ItemIndex)) And Not IncludeEmptyItems) Then
                            retVal(retValIndex, 1) = allItems(ItemIndex): added = True
                        End If
                        If added Then retVal(retValIndex, 0) = Empty
                    End If
                    If added Then retValIndex = retValIndex + 1
                End If
            Next
        End If
    Else
        ' one-dimensional array with keys
        ReDim retVal(0 To (retValCount - 1))
        If SortedByKeys Then
            ' sorted by keys, no items, run over all keys and add items/keys that match criteria
            For KeyIndex = 0 To (mKeyCount - 1)
                added = False
                ItemIndex = allKeyItemIndexes(KeyIndex)
                If IsObject(allItems(ItemIndex)) Then
                    added = True
                ElseIf IncludeEmptyItems Or (Not IsEmpty(allItems(ItemIndex)) And Not IncludeEmptyItems) Then
                    added = True
                End If
                If added Then
                    ' set key
                    If mEmulateScriptingDictionary Then
                        If isStoredKeyOfTypeObject(allKeys(KeyIndex)) Then
                            Set retVal(retValIndex) = convertDictCollectionToScriptingDictionaryKey(allKeys(KeyIndex))
                        Else
                            retVal(retValIndex) = convertDictCollectionToScriptingDictionaryKey(allKeys(KeyIndex))
                        End If
                    ElseIf mEmulateCollection Then
                        retVal(retValIndex) = convertDictCollectionToCollectionKey(allKeys(KeyIndex))
                    Else
                        retVal(retValIndex) = allKeys(KeyIndex)
                    End If
                    retValIndex = retValIndex + 1
                End If
            Next
            ' add items without keys at end
            If IncludeEmptyKeys Then
                ' add items without keys
                For ItemIndex = 0 To (mItemCount - 1)
                    added = False
                    If allItemKeyIndexes(ItemIndex) = -1 Then
                        If IsObject(allItems(ItemIndex)) Then
                            added = True
                        ElseIf IncludeEmptyItems Or (Not IsEmpty(allItems(ItemIndex)) And Not IncludeEmptyItems) Then
                            added = True
                        End If
                    End If
                    If added Then
                        retVal(retValIndex) = Empty
                        retValIndex = retValIndex + 1
                    End If
                Next
            End If
            
        Else
            ' no sorting (order of items), no items
            ' include items, no sorting, run over items and add items/keys that match criteria
            For ItemIndex = 0 To (mItemCount - 1)
                added = False
                KeyIndex = allItemKeyIndexes(ItemIndex)
                If IncludeEmptyKeys Or (Not IncludeEmptyKeys And KeyIndex > -1) Then
                    If KeyIndex > -1 Then
                        ' key exists
                        If IsObject(allItems(ItemIndex)) Then
                            added = True
                        ElseIf IncludeEmptyItems Or (Not IsEmpty(allItems(ItemIndex)) And Not IncludeEmptyItems) Then
                            added = True
                        End If
                        If added Then
                            ' set key
                            If mEmulateScriptingDictionary Then
                                If isStoredKeyOfTypeObject(allKeys(KeyIndex)) Then
                                    Set retVal(retValIndex) = convertDictCollectionToScriptingDictionaryKey(allKeys(KeyIndex))
                                Else
                                    retVal(retValIndex) = convertDictCollectionToScriptingDictionaryKey(allKeys(KeyIndex))
                                End If
                            ElseIf mEmulateCollection Then
                                retVal(retValIndex) = convertDictCollectionToCollectionKey(allKeys(KeyIndex))
                            Else
                                retVal(retValIndex) = allKeys(KeyIndex)
                            End If
                        End If
                    Else
                        ' key does not exist
                        retVal(retValIndex) = Empty: added = True
                    End If
                    If added Then retValIndex = retValIndex + 1
                End If
            Next
        End If
    End If

ExitFunction:
    internalGetKeys = retVal
End Function

' Returns all items (in the order they where added) as variant array or an empty array if there are no items
' Variant array is compatible with Scripting.Dictionary
Public Property Get Items(Optional IncludeEmptyItems As Boolean = True, Optional IncludeItemsWithoutKeys As Boolean = True) As Variant
    Items = internalGetItems(IncludeEmptyItems, IncludeItemsWithoutKeys, False, False)
End Property

' UNTESTED
' Returns all items (in the order they where added) and keys as two-dimensional Variant array with array(ItemIndex,0) = Item and array(ItemIndex,1) = Key
' Returns a one-dimensional empty array if DictCollection contains no items
Public Property Get ItemsAndKeys(Optional IncludeEmptyItems As Boolean = True, Optional IncludeItemsWithoutKeys As Boolean = True) As Variant
    ItemsAndKeys = internalGetItems(IncludeEmptyItems, IncludeItemsWithoutKeys, False, True)
End Property

' UNTESTED
' Returns all items sorted by their keys as Variant array
' Returns a one-dimensional empty array if DictCollection contains no items
Public Property Get ItemsSortedByKeys(Optional IncludeEmptyItems As Boolean = True, Optional IncludeItemsWithoutKeys As Boolean = False) As Variant
    ItemsSortedByKeys = internalGetItems(IncludeEmptyItems, IncludeItemsWithoutKeys, True, False)
End Property

' UNTESTED
' Returns all items (sorted by their keys) and their keys as two-dimensional Variant array with array(ItemIndex,0) = Item and array(ItemIndex,1) = Key
' Returns a one-dimensional empty array if DictCollection contains no items
Public Property Get ItemsAndKeysSortedByKeys(Optional IncludeEmptyItems As Boolean = False, Optional IncludeItemsWithoutKeys As Boolean = False) As Variant
    ItemsAndKeysSortedByKeys = internalGetItems(IncludeEmptyItems, IncludeItemsWithoutKeys, True, True)
End Property

' internal function that builds one/two-dimensional output array
Private Function internalGetItems(Optional IncludeEmptyItems As Boolean = True, Optional IncludeItemsWithoutKeys As Boolean = True, Optional SortedByKeys As Boolean = False, Optional IncludeKeys As Boolean = False) As Variant
    Dim retVal() As Variant, retValIndex As Long, retValCount As Long, ItemIndex As Long, KeyIndex As Long, added As Boolean
    If mItemCount < 1 Then retVal = Array(): GoTo ExitFunction
    
    ' sort keys if needed
    If (IncludeKeys Or SortedByKeys) And mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
    
    retValCount = 0
    If IncludeEmptyItems Then
        retValCount = mItemCount
    Else
        ' count non-empty items
        For ItemIndex = 0 To (mItemCount - 1)
            If IsObject(allItems(ItemIndex)) Then
                retValCount = retValCount + 1
            ElseIf Not IsEmpty(allItems(ItemIndex)) Then
                retValCount = retValCount + 1
            End If
        Next
    End If
    
    retValIndex = 0
    If IncludeKeys Then
        ReDim retVal(0 To (retValCount - 1), 1)
        If SortedByKeys Then
            ' include keys, sorted by keys
            For KeyIndex = 0 To (mKeyCount - 1)
                added = False
                ItemIndex = allKeyItemIndexes(KeyIndex)
                If IsObject(allItems(ItemIndex)) Then
                    Set retVal(retValIndex, 0) = allItems(ItemIndex): added = True
                ElseIf IncludeEmptyItems Or (Not IsEmpty(allItems(ItemIndex)) And Not IncludeEmptyItems) Then
                    retVal(retValIndex, 0) = allItems(ItemIndex): added = True
                End If
                If added Then
                    ' set key
                    If mEmulateScriptingDictionary Then
                        If isStoredKeyOfTypeObject(allKeys(KeyIndex)) Then
                            Set retVal(retValIndex, 1) = convertDictCollectionToScriptingDictionaryKey(allKeys(KeyIndex))
                        Else
                            retVal(retValIndex, 1) = convertDictCollectionToScriptingDictionaryKey(allKeys(KeyIndex))
                        End If
                    ElseIf mEmulateCollection Then
                        retVal(retValIndex, 1) = convertDictCollectionToCollectionKey(allKeys(KeyIndex))
                    Else
                        retVal(retValIndex, 1) = allKeys(KeyIndex)
                    End If
                    retValIndex = retValIndex + 1
                End If
            Next
        Else
            ' include keys, no sorting
            For ItemIndex = 0 To (mItemCount - 1)
                added = False
                If IsObject(allItems(ItemIndex)) Then
                    Set retVal(retValIndex, 0) = allItems(ItemIndex): added = True
                ElseIf IncludeEmptyItems Or (Not IsEmpty(allItems(ItemIndex)) And Not IncludeEmptyItems) Then
                    retVal(retValIndex, 0) = allItems(ItemIndex): added = True
                End If
                If added Then
                    KeyIndex = allItemKeyIndexes(ItemIndex)
                    If KeyIndex > -1 Then
                        ' set key
                        If mEmulateScriptingDictionary Then
                            If isStoredKeyOfTypeObject(allKeys(KeyIndex)) Then
                                Set retVal(retValIndex, 1) = convertDictCollectionToScriptingDictionaryKey(allKeys(KeyIndex))
                            Else
                                retVal(retValIndex, 1) = convertDictCollectionToScriptingDictionaryKey(allKeys(KeyIndex))
                            End If
                        ElseIf mEmulateCollection Then
                            retVal(retValIndex, 1) = convertDictCollectionToCollectionKey(allKeys(KeyIndex))
                        Else
                            retVal(retValIndex, 1) = allKeys(KeyIndex)
                        End If
                    Else
                        retVal(retValIndex, 1) = Empty
                    End If
                    retValIndex = retValIndex + 1
                End If
            Next
        End If
    Else
        ' do not include keys
        ReDim retVal(0 To (retValCount - 1))
        If SortedByKeys Then
            ' no keys, sorted by keys
            For KeyIndex = 0 To (mKeyCount - 1)
                added = False
                ItemIndex = allKeyItemIndexes(KeyIndex)
                If IsObject(allItems(ItemIndex)) Then
                    Set retVal(retValIndex) = allItems(ItemIndex): retValIndex = retValIndex + 1
                ElseIf IncludeEmptyItems Or (Not IsEmpty(allItems(ItemIndex)) And Not IncludeEmptyItems) Then
                    retVal(retValIndex) = allItems(ItemIndex): retValIndex = retValIndex + 1
                End If
            Next
        Else
            ' no sorting
            For ItemIndex = 0 To (mItemCount - 1)
                If IsObject(allItems(ItemIndex)) Then
                    Set retVal(retValIndex) = allItems(ItemIndex): retValIndex = retValIndex + 1
                ElseIf IncludeEmptyItems Or (Not IsEmpty(allItems(ItemIndex)) And Not IncludeEmptyItems) Then
                    retVal(retValIndex) = allItems(ItemIndex): retValIndex = retValIndex + 1
                End If
            Next
        End If
    End If
    
ExitFunction:
    internalGetItems = retVal
End Function

' returns the items key index from the internal key array or -1 if the item has no key or the ItemIndex does not exist (with .ThrowErrors=False)
' the key index is always zero-based
Public Property Get KeyIndexOfItemAt(Index As Long) As Long
    Dim index2 As Long
    index2 = Index - mExternalIndexOffset
    If mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
    If index2 > -1 And index2 <= (mItemCount - 1) Then
        KeyIndexOfItemAt = allItemKeyIndexes(index2)
    Else
        If mThrowErrors Then Err.Raise 9, "DictCollection", "No key found at index " & Index
        KeyIndexOfItemAt = -1
    End If
End Property

' returns the key from the internal key array at KeyIndex or "" if KeyIndex does not exist
' KeyIndex is always zero-based
Public Property Get KeyAtKeyIndexAsString(KeyIndex As Long) As String
    KeyAtKeyIndexAsString = ""
    If mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
    If KeyIndex > -1 And KeyIndex <= (mKeyCount - 1) Then
        KeyAtKeyIndexAsString = allKeys(KeyIndex)
    Else
        If mThrowErrors Then Err.Raise 9, "DictCollection", "No key found at index " & KeyIndex
    End If
End Property

' returns the key from the internal key array at KeyIndex or Empty if KeyIndex does not exist
' KeyIndex is always zero-based
Public Property Get KeyAtKeyIndex(KeyIndex As Long) As Variant
    KeyAtKeyIndex = Empty
    If mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
    If KeyIndex > -1 And KeyIndex <= (mKeyCount - 1) Then
        If mEmulateScriptingDictionary Then
            If isStoredKeyOfTypeObject(allKeys(KeyIndex)) Then
                Set KeyAtKeyIndex = convertDictCollectionToScriptingDictionaryKey(allKeys(KeyIndex))
            Else
                KeyAtKeyIndex = convertDictCollectionToScriptingDictionaryKey(allKeys(KeyIndex))
            End If
        ElseIf EmulateCollection Then
            KeyAtKeyIndex = convertDictCollectionToCollectionKey(allKeys(KeyIndex))
        Else
            KeyAtKeyIndex = allKeys(KeyIndex)
        End If
    Else
        If mThrowErrors Then Err.Raise 9, "DictCollection", "No key found at key index " & KeyIndex
    End If
End Property

' UNTESTED
' returns the index of key in the internal key array that is greater than SearchedKey
' returns -1 if internal key array is empty or SearchedKey could not be converted to internal String representation
' returned index is always zero-based
Public Function FindKeyInsertIndex(SearchedKey As Variant, Optional CompareMode As VbCompareMethod = vbBinaryCompare) As Long
    Dim result As EvaluateKeyResult
    If mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
    If mEmulateScriptingDictionary Then
        result = evaluateKey(convertScriptingDictionaryToDictCollectionKey(SearchedKey))
    ElseIf mEmulateCollection Then
        result = evaluateKey(convertCollectionToDictCollectionKey(SearchedKey))
    Else
        result = evaluateKey(SearchedKey)
    End If
    If result.WasInvalid Or mKeyCount < 1 Then
        FindKeyInsertIndex = -1
    Else
        FindKeyInsertIndex = internalFindKeyIndex(result.key, CompareMode).Index
    End If
End Function

' UNTESTED
' returns all keys that start with SearchText as variant array containing the internal String representation of the keys
Public Function FindKeysThatStartWith(SearchText As String, Optional CompareMode As VbCompareMethod = vbBinaryCompare) As Variant
    Dim retVal As Variant, result As FindKeyResult, i As Long
    retVal = Array()
    If mKeyCount < 1 Then GoTo ExitFunction
    If mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
    result = internalFindKeyIndex(SearchText, CompareMode)
    If result.Index <= (mKeyCount - 1) Then
        ' go downwards and add all matching keys to array, stop on first key that does not match
        For i = result.Index To 0 Step -1
            If UtilStringStartsWith(allKeys(i), SearchText, CompareMode) Then UtilAddArrayValue retVal, allKeys(i) Else Exit For
        Next
        ' go upwards and add all matching keys to array, stop on first key that does not match
        For i = result.Index + 1 To (mKeyCount - 1)
            If UtilStringStartsWith(allKeys(i), SearchText, CompareMode) Then UtilAddArrayValue retVal, allKeys(i) Else Exit For
        Next
    Else
        ' SearchText is greater than the last key, add last key if it matches
        If UtilStringStartsWith(allKeys(mKeyCount - 1), SearchText, CompareMode) Then UtilAddArrayValue retVal, allKeys(mKeyCount - 1)
    End If
ExitFunction:
    FindKeysThatStartWith = retVal
End Function

' UNTESTED
' adds keys with specified value (empty by default)
Public Function AddKeys(Arr As Variant, Optional Value As Variant = Empty) As DictCollection
    Dim i As Long, dimensions As Integer
    Set AddKeys = Me
    dimensions = UtilArrayDimensions(Arr)
    Select Case dimensions
        Case 0: ' Unitialized Array
            If mThrowErrors Then Err.Raise 13, "DictCollection", "Can't add from uninitialized array. Must be 1-dimensional array."
            Exit Function
        Case 2: ' Two-Dimensional Array
            If mThrowErrors Then Err.Raise 13, "DictCollection", "Can't add from 2-dimensional array. Must be 1-dimensional array."
            Exit Function
        Case -1:
            ' One-Dimensional Nested Array
            If mThrowErrors Then Err.Raise 13, "DictCollection", "Can't add from 1-dimensional nested array. Must be 1-dimensional array."
            Exit Function
        Case 1:
            ' One-Dimensional Array
            For i = LBound(Arr) To UBound(Arr)
                Me.Add Arr(i), Value
            Next
    End Select
End Function


' adds key/value pairs from a two-dimensional or nested array, chainable
Public Function AddPairs(Arr As Variant, Optional KeyIsFirstColumn = True) As DictCollection
    Dim i As Long, KeyIndex As Long, valIndex As Long, dimensions As Integer
    Set AddPairs = Me
    dimensions = UtilArrayDimensions(Arr)
    Select Case dimensions
        Case 0: ' Unitialized Array
            If mThrowErrors Then Err.Raise 13, "DictCollection", "Can't add from uninitialized array. Must be 2-dimensional or nested array."
            Exit Function
        Case 1: ' One-Dimensional Array
            If mThrowErrors Then Err.Raise 13, "DictCollection", "Can't add from 1-dimensional array. Must be 2-dimensional or nested array."
            Exit Function
        Case -1:
            ' One-Dimensional Nested Array
            For i = LBound(Arr) To UBound(Arr)
                If KeyIsFirstColumn Then
                    ' key is first index and value is second
                    KeyIndex = LBound(Arr(i)): valIndex = KeyIndex + 1
                Else
                    ' the other way round
                    valIndex = LBound(Arr(i)): KeyIndex = valIndex + 1
                End If
                ' do not add empty keys unless we are emulating Scripting.Dictionary
                If UBound(Arr(i)) > 0 Then
                    If Not IsEmpty(Arr(i)(KeyIndex)) Or mEmulateScriptingDictionary Then
                        Me.Add Arr(i)(KeyIndex), Arr(i)(valIndex)
                    Else
                        If mThrowErrors Then Err.Raise 13, "DictCollection", "Nested arr(" & i & ") contains empty key!"
                        ' otherwise ignore and continue
                    End If
                Else
                    If mThrowErrors Then Err.Raise 9, "DictCollection", "Nested arr(" & i & ") must have 2 items but contains only 1."
                    ' otherwise ignore and continue
                End If
            Next
        Case 2:
            ' Two-Dimensional Array
            If UBound(Arr, 2) > 0 Then
                If KeyIsFirstColumn Then
                    ' key is first index and value is second
                    KeyIndex = LBound(Arr, 2): valIndex = KeyIndex + 1
                Else
                    ' the other way round
                    valIndex = LBound(Arr, 2): KeyIndex = valIndex + 1
                End If
                For i = LBound(Arr, 1) To UBound(Arr, 1)
                    ' do not add empty keys unless we are emulating Scripting.Dictionary
                    If Not IsEmpty(Arr(i, KeyIndex)) Or mEmulateScriptingDictionary Then
                        Me.Add Arr(i, KeyIndex), Arr(i, valIndex)
                    Else
                        If mThrowErrors Then Err.Raise 13, "DictCollection", "2-dimensional arr(" & i & ") contains empty key!"
                        ' otherwise ignore and continue
                    End If
                Next
            Else
                If mThrowErrors Then Err.Raise 9, "DictCollection", "Second dimension of 2-dimensional array must be > 0!"
                Exit Function
            End If
    End Select
End Function

' Creates a new DictCollection from an existing one by dropping all keys that match any key stored in KeyArr.
' Does not recoursivey traverse the DictCollection. Use .Flatten() -> .Filter() -> .Unflatten() if you want to drop SubCollection keys.
' Ignores non-string and empty keys (not 100% compatible with .EmulateDictionary=True).
' If you have only one string, use VBA.Array(myString) to create the KeyArr
' AllowWildcards=False: only exact matches
' AllowWildcards=True:  inexact matches using the asterisk '*' character.
'                       Examples: 1) '*test*' -> drops keys containing 'test' 2) '*abc' -> drops keys ending with 'abc' 3) '_*' -> drops keys starting with '_'
' CompareMode:          toggles between case sensitve (vbBinaryCompare, default) key comparison and case insensitive (vbTextCompare)
' Example usage:        Dim dc As New DictCollection: dc.AddPairs Array(Array("a", 1), Array("b", 2), Array("bx", 3))
'                       Dim dc2 As DictCollection: Set dc2 = dc.DropKeysToNew(Array("*b*"), True)
'                       '-> will create new DC and drop any keys that contain, start or end with "b" -> drops item 2 and 3
Public Function DropKeysToNew(KeyArr As Variant, Optional AllowWildcards As Boolean = False, Optional CompareMode As VbCompareMethod = vbBinaryCompare) As DictCollection
    Set DropKeysToNew = internalFilterKeysToNew(KeyArr, AllowWildcards, CompareMode, 1, mNonExistingValue)
End Function

' Creates a new DictCollection from an existing one by keeping all keys that match any key stored in KeyArr.
' Does not recoursivey traverse the DictCollection. Use .Flatten() -> .Filter() -> .Unflatten() if you want to keep SubCollection keys.
' Ignores non-string and empty keys (not 100% compatible with .EmulateDictionary=True).
' AllowWildCards=False: only exact matches
' AllowWildCards=True:  inexact matches using the asterisk '*' character.
'                       Examples: 1) '*test*' -> keeps keys containing 'test' 2) '*abc' -> keeps keys ending with 'abc' 3) '_*' -> keeps keys starting with '_'
' CompareMode:          toggles between case sensitve (vbBinaryCompare, default) key comparison and case insensitive (vbTextCompare)
' Example usage:        Dim dc As New DictCollection: dc.AddPairs Array(Array("a", 1), Array("b", 2), Array("bx", 3))
'                       Dim dc2 As DictCollection: Set dc2 = dc.KeepKeysToNew(Array("b"))
'                       '-> will keep any key that matches "b" -> keeps item 2
'                       Set dc = New DictCollection: dc.AddPairs Array(Array("a", 1), Array("b", 2), Array("bx", 3))
'                       Set dc2 = dc.KeepKeysToNew(Array("*B*"), True, vbTextCompare)
'                       '-> will keep any item that contains, starts or ends with "B" or "b" -> keeps item 2 and 3
Public Function KeepKeysToNew(KeyArr As Variant, Optional AllowWildcards As Boolean = False, Optional CompareMode As VbCompareMethod = vbBinaryCompare) As DictCollection
    Set KeepKeysToNew = internalFilterKeysToNew(KeyArr, AllowWildcards, CompareMode, 2, mNonExistingValue)
End Function

' Creates a new DictCollection with keys and items that match the keys specified in KeyArr.
' The items will have the same order as the specified keys in KeyArr. Unmatched keys in KeyArr will be included with NotFoundValue as items.
' Does not recoursivey traverse the DictCollection. Use .Flatten() -> .Filter() -> .Unflatten() if you want to keep SubCollection keys.
' Ignores non-string and empty keys (not 100% compatible with .EmulateDictionary=True).
' CompareMode:          toggles between case sensitve (vbBinaryCompare, default) key comparison and case insensitive (vbTextCompare)
' Example usage:        Dim dc As New DictCollection: dc.AddPairs Array(Array("a", 1), Array("b", 2), Array("c", 3))
'                       Dim dc2 As DictCollection: Set dc2 = dc.MatchKeysToNew(Array("a", "b", "x"), Empty)
'                       '-> will return DC with items 1 and 2 and an Empty item for key "x"
Public Function MatchKeysToNew(KeyArr As Variant, Optional NotFoundValue As Variant = "[NONEXISTING]", Optional CompareMode As VbCompareMethod = vbBinaryCompare) As DictCollection
    Set MatchKeysToNew = internalFilterKeysToNew(KeyArr, False, CompareMode, 3, NotFoundValue)
End Function

' internal filter function that returns a new DictCollection according to the specified filter function
' action: 1 = drop, 2 = keep, 3 = match
Private Function internalFilterKeysToNew(KeyArr As Variant, AllowWildcards As Boolean, CompareMode As VbCompareMethod, action As Integer, NotFoundValue As Variant) As DictCollection
    Dim retVal As DictCollection, lower As Long, upper As Long, keyArr2 As Variant, tempDC As DictCollection
    Dim i As Long, j As Long, match As Boolean, currentKey As String, wildcardSplits As Variant
    Set retVal = New DictCollection
    DictCollection.CopyAllSettingsExceptCollectionType Me, retVal
    If Not IsArray(KeyArr) Then
        If mThrowErrors Then If mThrowErrors Then Err.Raise 13, "KeyArr is not an array"
    Else
        Select Case action
            ' drop and keep items
            Case 1, 2:
                ' build cleaned key array
                keyArr2 = Array()
                lower = LBound(KeyArr): upper = UBound(KeyArr)
                For i = lower To upper
                    If TypeName(KeyArr(i)) = "String" Then
                        If Len(KeyArr(i)) > 0 Then UtilAddArrayValue keyArr2, KeyArr(i)
                    End If
                Next
                If AllowWildcards And UBound(keyArr2) > -1 Then
                    ' produce splits
                    ReDim wildcardSplits(0 To UBound(keyArr2))
                    For j = 0 To UBound(keyArr2)
                        wildcardSplits(j) = Split(keyArr2(j), "*")
                    Next
                End If
                For i = 0 To mItemCount - 1
                    ' process only items that have key (drop others)
                    If allItemKeyIndexes(i) > -1 Then
                        currentKey = allKeys(allItemKeyIndexes(i))
                        match = False
                        For j = 0 To UBound(keyArr2)
                            If AllowWildcards Then
                                match = UtilStringMatchesWildcard(currentKey, wildcardSplits(j), CompareMode)
                            Else
                                match = (StrComp(currentKey, keyArr2(j), CompareMode) = 0)
                            End If
                            If match Then Exit For ' exit on first match
                        Next
                        If (action = 1 And Not match) Or (action = 2 And match) Then
                            If TypeName(allItems(i)) = "DictCollection" Then
                                ' copy key and clone DictCollection item
                                Set tempDC = allItems(i).CloneToNew
                                retVal.Add currentKey, tempDC
                            Else
                                ' copy item and key
                                retVal.Add currentKey, allItems(i)
                            End If
                        End If
                    End If
                Next
            ' match keys
            Case 3:
                Dim fkr As FindKeyResult
                ' run over all keys to match
                For i = 0 To UBound(KeyArr)
                    If TypeName(KeyArr(i)) = "String" Then
                        currentKey = KeyArr(i)
                        If Len(currentKey) > 0 Then
                            ' try to find key
                            fkr = internalFindKeyIndex(currentKey, CompareMode)
                            If fkr.Exact Then
                                ' key found, add item (clone if DictCollection)
                                If TypeName(allItems(allKeyItemIndexes(fkr.Index))) = "DictCollection" Then
                                    Set tempDC = allItems(allKeyItemIndexes(fkr.Index)).CloneToNew
                                    retVal.Add currentKey, tempDC
                                Else
                                    retVal.Add currentKey, allItems(allKeyItemIndexes(fkr.Index))
                                End If
                            Else
                                ' key not found, add notFoundValue
                                retVal.Add currentKey, NotFoundValue
                            End If
                        Else
                            ' empty key, add notFoundValue
                            retVal.Add currentKey, NotFoundValue
                        End If
                    End If
                Next
        End Select
    End If
    Set internalFilterKeysToNew = retVal
End Function


' can be enabled for testing purposes
' Friend Property Get InternalItems() As Variant(): InternalItems = allItems: End Property
' Friend Property Get InternalItemKeyIndexes() As Long(): InternalItemKeys = allItemKeyIndexes: End Property
' Friend Property Get InternalKeys() As String(): InternalKeys = allKeys: End Property
' Friend Property Get InternalKeyItemIndexes() As Long(): InternalKeyItemIndexes = allKeyItemIndexes: End Property
' ======================== END: KEYS AND ITEMS, FINDKEY FUNCTIONS, TOARRAY ==========================================================================


' ======================== START: TYPESAFE ITEM ACCESS ==============================================================================================
' returns the item as DictCollection
' returns a nonexisting DictCollection with .ThrowErrors=False if the index or key was not found or the item is not a DictCollecttion or
' nonexisting DictCollections have .CollectionType=0 and a default value of "[NONEXISTING]" if not configured otherwise with .NonExistingValue
Public Property Get AsDC(IndexOrKey As Variant) As DictCollection
    Dim retVal As DictCollection, result As EvaluateIndexOrKeyResult, tname As String
    result = internalGetItem(IndexOrKey, False)
    If Not result.WasFound Then
        ' errors will be thrown by internalGetItem
        Set retVal = New DictCollection: CopySubCollectionChainingSettings Me, retVal: retVal.CollectionType = 0 ' nonexisting
    Else
        tname = TypeName(allItems(result.Index))
        If tname = "DictCollection" Then
            Set retVal = allItems(result.Index)
        Else
            If mThrowErrors Then Err.Raise 13, "Item at index or key is not a DictCollection"
            Set retVal = New DictCollection: CopySubCollectionChainingSettings Me, retVal: retVal.CollectionType = 0 ' nonexisting
        End If
    End If
    Set AsDC = retVal
End Property

' UNTESTED
' returns True if the item at Index or Key is a DictCollection
' returns False if the item at Index or Key is not a DictCollection or the Index or Key was not found (with .ThrowErrors=False)
Public Property Get ItemIsDC(IndexOrKey As Variant) As Boolean
    Dim result As EvaluateIndexOrKeyResult: result = internalGetItem(IndexOrKey)
    ItemAtIsDC = result.WasFound And result.IsObject
    If ItemAtIsDC Then ItemAtIsDC = (TypeName(allItems(result.Index)) = "DictCollection")
End Property

' UNTESTED
' returns True if the item at Index or Key is an Object
' returns False if the item at Index or Key is not an Object or the Index or Key was not found (with .ThrowErrors=False)
Public Property Get ItemIsObject(IndexOrKey As Variant) As Boolean
    Dim result As EvaluateIndexOrKeyResult: result = internalGetItem(IndexOrKey)
    ItemAtIsDC = result.WasFound And result.IsObject
End Property
' ======================== END: TYPESAFE ITEM ACCESS ================================================================================================


' ======================== START: INTERNAL HELPER FUNCTIONS - LOOKUP, SORTING, COPYING ==============================================================
' gets the item index from a Variant that can be both, index or key and remembers the source data type
Private Function evaluateIndexOrKey(IndexOrKey As Variant) As EvaluateIndexOrKeyResult
    Dim retVal As EvaluateIndexOrKeyResult
    retVal.Index = -1 ' default value
    Select Case TypeName(IndexOrKey)
        Case "String"
            retVal.WasKey = True
            If IndexOrKey = vbNullString Then
                'do not search for key
                retVal.WasFound = False
            Else
                'a string was passed, search for key
                Dim key As String, resultKey As FindKeyResult
                key = IndexOrKey
                resultKey = internalFindKeyIndex(key, mCompareMode)
                If resultKey.Exact Then
                    retVal.WasFound = True: retVal.Index = allKeyItemIndexes(resultKey.Index) 'key was found
                Else
                    retVal.WasFound = False 'key was not found
                End If
            End If
        Case "Integer", "Long", "Byte"
            'an integer was passed, emulate VBA.Collection behavior by accepting all these types
            retVal.WasIndex = True
            retVal.Index = IndexOrKey - mExternalIndexOffset
            retVal.WasFound = (retVal.Index > -1) And (retVal.Index <= (mItemCount - 1))
        Case "Single", "Double", "Currency", "Decimal", "Date"
            'a fractional number was passed, convert to long
            retVal.WasIndex = True
            retVal.Index = CLng(IndexOrKey) - mExternalIndexOffset
            retVal.WasFound = (retVal.Index > -1) And (retVal.Index <= (mItemCount - 1))
        Case Else
            'argument is neither number nor string
            retVal.WasFound = False
    End Select
    evaluateIndexOrKey = retVal
End Function

' converts a Variant key to String and remembers the source data type
Private Function evaluateKey(key As Variant) As EvaluateKeyResult
    Dim retVal As EvaluateKeyResult: retVal.WasInvalid = True
    Select Case TypeName(key)
        Case "Integer", "Long", "Single", "Double", "Currency", "Decimal", "Date": retVal.WasInvalid = False: retVal.WasNumber = True: retVal.key = CStr(key)
        Case "String": retVal.WasInvalid = False: retVal.WasString = True: retVal.key = key
        Case "Empty": retVal.WasInvalid = False: retVal.WasEmpty = True: retVal.key = ""
        Case "Error": If IsMissing(key) Then retVal.WasInvalid = False: retVal.WasEmpty = True: retVal.key = ""
    End Select
    evaluateKey = retVal
End Function

' Returns index of key. Snaps to first stored key that matches the searched key (e.g. case insensitive search).
' If key is not found, returns
' 1) [-1] if searched key is less than first stored key
' 2) [last stored key index + 1] if searched key is greater than last stored key
Private Function internalFindKeyIndex(key As String, CompareMode As VbCompareMethod) As FindKeyResult
    Dim upper As Long, lower As Long, row As Long, counter As Long, retVal As FindKeyResult, rowBefore As Long
    retVal.Exact = False
    If mItemCount < 1 Or mKeyCount < 1 Then
        retVal.Index = 0
    Else
        lower = 0:  upper = mKeyCount - 1: counter = 0:
        'compare with last key
        Select Case StrComp(key, allKeys(upper), CompareMode) 'compare with last element
            Case 0: retVal.Index = upper: retVal.Exact = True 'equals last key
            Case 1: retVal.Index = upper + 1 'greater than last key
            Case -1
                'compare with first key
                Select Case StrComp(key, allKeys(lower), CompareMode)
                    Case 0: retVal.Index = lower: retVal.Exact = True 'equals first key
                    Case -1: retVal.Index = lower 'less than first key
                    Case 1:
                        ' search using quick search
                        row = (upper - lower) \ 2
                        Do
                            counter = counter + 1
                            Select Case StrComp(key, allKeys(row), CompareMode)
                                Case 0:
                                    ' make it snap to row before first key that matches by going back until key before does not match
                                    rowBefore = row - 1
                                    Do While rowBefore > -1 And StrComp(key, allKeys(rowBefore), CompareMode) = 0
                                        row = rowBefore: rowBefore = row - 1
                                    Loop
                                    retVal.Index = row: retVal.Exact = True
                                Case 1: retVal.Index = row + 1: lower = row: row = lower + ((upper - lower) \ 2)
                                Case -1: retVal.Index = row: upper = row: row = lower + ((upper - lower) \ 2)
                            End Select
                        Loop While ((upper - lower) > 1) And (retVal.Exact = False) And counter < 7000
                        If counter = 7000 Then Debug.Print "!!!!!! Bug found in DictCollection: Endless loop at internalFindKeyIndex(""" & key & """)!!!!!!!"
                End Select
        End Select
    End If
    internalFindKeyIndex = retVal
End Function

' performs key sorting and removes duplicates can exist when adding items and keys while LazySorting=True
Private Sub sortKeysAndRemoveDuplicates()
    If mKeyCount < 0 Then Exit Sub
    sortKeys 0, (mKeyCount - 1)
    removeKeyDuplicates
End Sub

' removes all duplicate key entries by running over all sorted keys and dropping all but the last key/item pairs while preserving the first item index for that key
Private Sub removeKeyDuplicates()
    Dim i As Long, j As Long, k As Long, lastKey As String, itemIndexes As Variant, Offset As Long, match As Boolean, last As Boolean, swappedValue As Variant
    ' remove duplicate keys and corresponding items
    lastKey = allKeys(0): itemIndexes = VBA.Array(): i = 1
    While i <= UBound(allKeys)
        If StrComp(lastKey, allKeys(i), mCompareMode) = 0 Then
            match = True
            ' add item index to array
            If UBound(itemIndexes) = -1 Then
                ' add previous and current item indexes
                ReDim Preserve itemIndexes(0 To 1)
                itemIndexes(0) = allKeyItemIndexes(i - 1)
                itemIndexes(1) = allKeyItemIndexes(i)
            Else
                ' add current item index
                ReDim Preserve itemIndexes(0 To UBound(itemIndexes) + 1)
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
                ' set the key of that item to the key of the last added item (in case the keys are equal but the casing is different)
                allKeys(i - 1) = allKeys(allItemKeyIndexes(itemIndexes(UBound(itemIndexes))))
                ' remove all similar keys except last one
                copyKeysDown i - 1, UBound(allKeys), Offset, True
                ' set the overwriting items key to the remaining unique key
                allItemKeyIndexes(itemIndexes(0)) = (i - 1) - Offset
                
                ' set arrays to new size
                ReDim Preserve allKeys(0 To UBound(allKeys) - Offset)
                ReDim Preserve allKeyItemIndexes(0 To UBound(allKeyItemIndexes) - Offset)
                
                ' remove corresponding items, starting from last to second
                For j = UBound(itemIndexes) To 1 Step -1
                    ' move last index one down with each iteration to avoid adjusting keyItemIndexes of last items in array multiple times
                    copyItemsDown itemIndexes(j) + 1, UBound(allItems) - (UBound(itemIndexes) - j), 1, True
                Next
                
                
                ReDim Preserve allItems(0 To UBound(allItems) - UBound(itemIndexes))
                ReDim Preserve allItemKeyIndexes(0 To UBound(allItemKeyIndexes) - UBound(itemIndexes))
                
                ' reset variables
                itemIndexes = VBA.Array()
                i = i - Offset
            End If
            If Not last Then lastKey = allKeys(i)
        End If
        i = i + 1
    Wend
    mItemCount = UBound(allItems) + 1
    mKeyCount = UBound(allKeys) + 1
End Sub

' source: https://wellsr.com/vba/2018/excel/vba-quicksort-macro-to-sort-arrays-fast/
Private Sub sortKeys(FromIndex As Long, ToIndex As Long)
    ' StrComp(string1, string2) -> [string1 < string2] = -1, [string1 > string2] = 1, [string1 = string2] = 0
    Dim splitKey As String, swappedKey As String, swappedKeyItemIndex As Long, swappedItemKeyIndex As Long, lower As Long, upper As Long
    lower = FromIndex
    upper = ToIndex
    splitKey = allKeys((FromIndex + ToIndex) \ 2)
    While (lower <= upper) 'divide
        While (StrComp(allKeys(lower), splitKey, mCompareMode) = -1 And lower < ToIndex)
           lower = lower + 1
        Wend
        While (StrComp(splitKey, allKeys(upper), mCompareMode) = -1 And upper > FromIndex)
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
    If (FromIndex < upper) Then sortKeys FromIndex, upper 'conquer
    If (lower < ToIndex) Then sortKeys lower, ToIndex 'conquer
End Sub

' copies items to higher index by offset and increases their key-to-item pointers if needed (to reflect the new item indexes)
' will copy only the items that can be copied without changing the array size
Private Sub copyItemsUp(FirstIndex As Long, LastIndex As Long, Offset As Long, IncreaseTheirKeyItemIndexByOffset As Boolean)
    Dim sourceIndex As Long, KeyIndex As Long, lower As Long, upper As Long
    ' ensure that copying is possible within array bounds
    If Offset >= mItemCount Then Exit Sub
    If FirstIndex < 0 Then lower = 0 Else lower = FirstIndex
    If LastIndex > (mItemCount - 1 - Offset) Then upper = mItemCount - 1 - Offset Else upper = LastIndex
    For sourceIndex = upper To lower Step -1
        If IsObject(allItems(sourceIndex)) Then Set allItems(sourceIndex + Offset) = allItems(sourceIndex) Else allItems(sourceIndex + Offset) = allItems(sourceIndex)
        KeyIndex = allItemKeyIndexes(sourceIndex)
        allItemKeyIndexes(sourceIndex + Offset) = KeyIndex
        If IncreaseTheirKeyItemIndexByOffset Then
            If KeyIndex > -1 And KeyIndex < mKeyCount Then allKeyItemIndexes(KeyIndex) = allKeyItemIndexes(KeyIndex) + Offset
        End If
    Next
End Sub

' copies items to lower index by offset and decreases their key-to-item pointers if needed (to reflect the new item indexes)
' will copy only the items that can be copied without changing the array size
Private Sub copyItemsDown(FirstIndex As Long, LastIndex As Long, Offset As Long, DecreaseTheirKeyItemIndexByOffset As Boolean)
    Dim sourceIndex As Long, KeyIndex As Long, lower As Long, upper As Long
    ' ensure that copying is possible within array bounds
    If Offset >= mItemCount Then Exit Sub
    If (FirstIndex - Offset) < 0 Then lower = Offset Else lower = FirstIndex
    If LastIndex >= mItemCount Then upper = mItemCount - 1 Else upper = LastIndex
    For sourceIndex = lower To upper Step 1
        If IsObject(allItems(sourceIndex)) Then Set allItems(sourceIndex - Offset) = allItems(sourceIndex) Else allItems(sourceIndex - Offset) = allItems(sourceIndex)
        KeyIndex = allItemKeyIndexes(sourceIndex)
        allItemKeyIndexes(sourceIndex - Offset) = KeyIndex
        If DecreaseTheirKeyItemIndexByOffset Then
            If KeyIndex > -1 And KeyIndex < mKeyCount Then allKeyItemIndexes(KeyIndex) = allKeyItemIndexes(KeyIndex) - Offset
        End If
    Next
End Sub

' copies keys to lower index by offset and increases their item-to-key pointers if needed (to reflect the new key indexes)
' will copy only the keys that can be copied without changing the array size
Private Sub copyKeysUp(FirstIndex As Long, LastIndex As Long, Offset As Long, IncreaseTheirItemKeyIndexByOffset As Boolean)
    Dim sourceIndex As Long, ItemIndex As Long, lower As Long, upper As Long
    ' ensure that copying is possible within array bounds
    If Offset >= mKeyCount Then Exit Sub
    If FirstIndex < 0 Then lower = 0 Else lower = FirstIndex
    If LastIndex > (mKeyCount - 1 - Offset) Then upper = mKeyCount - 1 - Offset Else upper = LastIndex
    For sourceIndex = upper To lower Step -1
        allKeys(sourceIndex + Offset) = allKeys(sourceIndex)
        ItemIndex = allKeyItemIndexes(sourceIndex)
        allKeyItemIndexes(sourceIndex + Offset) = ItemIndex
        If IncreaseTheirItemKeyIndexByOffset Then
            If ItemIndex > -1 And ItemIndex < mItemCount Then allItemKeyIndexes(ItemIndex) = allItemKeyIndexes(ItemIndex) + Offset
        End If
    Next
End Sub

' copies keys to higher index by offset and decreases their item-to-key pointers if needed (to reflect the new key indexes)
' will copy only the keys that can be copied without changing the array size
Private Sub copyKeysDown(FirstIndex As Long, LastIndex As Long, Offset As Long, DecreaseTheirItemKeyIndexByOffset As Boolean)
    Dim sourceIndex As Long, ItemIndex As Long, lower As Long, upper As Long
    ' ensure that copying is possible within array bounds
    If Offset >= mKeyCount Then Exit Sub
    lower = IIf((FirstIndex - Offset) < 0, Offset, FirstIndex)
    upper = IIf(LastIndex >= mKeyCount, mKeyCount - 1, LastIndex)
    For sourceIndex = lower To upper Step 1
        allKeys(sourceIndex - Offset) = allKeys(sourceIndex)
        ItemIndex = allKeyItemIndexes(sourceIndex)
        allKeyItemIndexes(sourceIndex - Offset) = ItemIndex
        If DecreaseTheirItemKeyIndexByOffset Then
            If ItemIndex > -1 And ItemIndex < mItemCount Then allItemKeyIndexes(ItemIndex) = allItemKeyIndexes(ItemIndex) - Offset
        End If
    Next
End Sub
' ======================== END: INTERNAL HELPER FUNCTIONS - LOOKUP, SORTING, COPYING ================================================================



' ======================== START: EMULATION HELPER FUNCTIONS =========================================================================================
' used by Scripting.Dictionary compatibility: converts Dictionary Variant keys to String keys that can be stored by DictCollection
Function convertScriptingDictionaryToDictCollectionKey(IndexOrKey As Variant) As String
    Dim key As String, tname As String
    tname = TypeName(IndexOrKey)
    Select Case tname
        Case "String":
            If IndexOrKey = "" Then key = "[EMP:]" Else key = IndexOrKey
        Case "Integer", "Long", "Single", "Double", "Currency", "Decimal", "Date":
            key = "[NUM:" & CStr(IndexOrKey) & "]"
        Case "Byte":
            Dim byteArray() As Byte: byteArray = IndexOrKey
            key = "[BTE:" & StrConv(byteArray, vbUnicode) & "]"
        Case "Boolean": key = "[BOL:" & CStr(IndexOrKey) & "]"
        Case "Error":
            Dim ErrNumber As String: ErrNumber = CStr(IndexOrKey)
            ' remove the "Error " part
            key = "[ERR:" & Mid(ErrNumber, 7, Len(ErrNumber) - 6) & "]"
        Case "Null": key = "[NUL:]"
        Case "Empty": key = "[EMP:]"
        Case Else:
            If IsObject(IndexOrKey) Then
                ' find object in array. if not found, set key to index after last index
                Dim Index As Long: Index = UtilFindArrayIndex(mScriptingDictionaryObjectKeys, IndexOrKey)
                If Index > -1 Then key = "[OBJ:" & Index & "]" Else key = "[OBJ:" & UBound(mScriptingDictionaryObjectKeys) + 1 & "]"
            End If
    End Select
    convertScriptingDictionaryToDictCollectionKey = key
End Function

' used by Scripting.Dictionary compatibility: converts a stored String key back to Scripting.Dictionary Variant (String, Number, Object, Error or Empty)
Private Function convertDictCollectionToScriptingDictionaryKey(key As String) As Variant
    Dim keyType As String, keyVal As String, originalDataType As String, retVal As Variant
    retVal = key
    If Len(key) >= 6 Then
        keyType = Mid(key, 2, 3): keyVal = Mid(key, 6, Len(key) - 6)
        Select Case keyType
            Case "NUM":
                originalDataType = mScriptingDictionaryStoredKeyTypes(key)
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
                Dim objectKeyIndex As Long: objectKeyIndex = getScriptingDictionaryObjectKeyIndex(key)
                If objectKeyIndex > -1 And objectKeyIndex <= UBound(mScriptingDictionaryObjectKeys) Then
                    Set retVal = mScriptingDictionaryObjectKeys(objectKeyIndex)
                Else
                    ' we get here if someone added something like "[OBJ:0]" as string key
                    Set retVal = Nothing
                End If
            Case "ERR": If keyVal = "448" Then retVal = UtilGetMissingValue Else retVal = CVErr(CLng(keyVal))
            Case "EMP"
                originalDataType = mScriptingDictionaryStoredKeyTypes(key)
                Select Case originalDataType
                    Case "String": retVal = ""
                    Case "Empty": retVal = Empty
                End Select
            Case "NUL": retVal = Null
        End Select
    End If
    If IsObject(retVal) Then Set convertDictCollectionToScriptingDictionaryKey = retVal Else convertDictCollectionToScriptingDictionaryKey = retVal
End Function

' used by Scripting.Dictionary compatibility: returns true if stored key is an object (needed to distinguish between set/let syntax)
Private Function isStoredKeyOfTypeObject(key As String) As Boolean
    isStoredKeyOfTypeObject = False
    If Len(key) >= 6 Then
        If Mid(key, 2, 3) = "OBJ" And Right(key, 1) = "]" Then isStoredKeyOfTypeObject = True
    End If
End Function

' used by Scripting.Dictionary compatibility: checks if stored key has a data type that is stored mScriptingDictionaryStoredKeyTypes:
' Integer, Long, Single, Double, Currency, Decimal, etc.
Private Function isScriptingDictionaryStoredKeyType(key As String) As Boolean
    isScriptingDictionaryStoredKeyType = False
    If Len(key) > 4 Then
        Select Case Left(key, 4)
            Case "[NUM", "[EMP":
                isScriptingDictionaryStoredKeyType = True
        End Select
    End If
End Function

' used by Scripting.Dictionary compatibility: stores the key object or the original number datatype if necessary
Private Sub storeScriptingDictionaryKeyInformation(ExternalKey As Variant, InternalKey As String)
     Dim objectKeyIndex As Long
     If TypeName(ExternalKey) <> "String" Then
         ' check if key was object and add it if it does not exist or re-set it, if it does exist
         objectKeyIndex = getScriptingDictionaryObjectKeyIndex(InternalKey)
         If UBound(mScriptingDictionaryObjectKeys) > -1 Or objectKeyIndex > -1 Then
             If objectKeyIndex > UBound(mScriptingDictionaryObjectKeys) Then UtilAddArrayValue mScriptingDictionaryObjectKeys, ExternalKey
        End If
    End If
    ' check if key needs to be stored and if yes, store original key type name for backward conversion
    If isScriptingDictionaryStoredKeyType(InternalKey) Then mScriptingDictionaryStoredKeyTypes(InternalKey) = TypeName(ExternalKey)
End Sub

' used by Scripting.Dictionary compatibility: checks if key is object and if yes, extracts index from key. if not, returns -1
Private Function getScriptingDictionaryObjectKeyIndex(key As String) As Long
    Dim retVal As Long: retVal = -1
    If Len(key) > 5 Then
        If Left(key, 5) = "[OBJ:" Then retVal = CLng(Mid(key, 6, Len(key) - 6))
    End If
    getScriptingDictionaryObjectKeyIndex = retVal
End Function

' used by Scripting.Dictionary compatibility: removes an object from the object keys array and decreases all object key references greater/equal index in allKeys by one
Private Sub removeScriptingDictionaryObjectKey(Index As Long)
    Dim i As Long, objectKeyIndex As Long
    ' remove old object key
    UtilRemoveArrayValueByIndex mScriptingDictionaryObjectKeys, Index
    ' run over all keys and decrease the existing object key indexes by 1 if they are greater than the old object key index
    For i = 0 To (mKeyCount - 1)
        objectKeyIndex = getScriptingDictionaryObjectKeyIndex(allKeys(i))
        If objectKeyIndex >= Index Then allKeys(i) = "[OBJ:" & objectKeyIndex - 1 & "]"
    Next
End Sub

' used by Collection compatibility: maps Collection Keys (Zero Length Strings and Missing) to DictCollection compatibe IndexOrKey
Function convertCollectionToDictCollectionKey(IndexOrKey As Variant) As Variant
    Dim key2 As Variant, tname As String
    If IsMissing(IndexOrKey) Then
        key2 = ""
    Else
        tname = TypeName(IndexOrKey)
        Select Case tname
            Case "String":
                If IndexOrKey = "" Then key2 = "[EMP:]" Else key2 = IndexOrKey
            Case Else:
                If IsObject(IndexOrKey) Then Set key2 = IndexOrKey Else key2 = IndexOrKey
        End Select
    End If
    If IsObject(key2) Then Set convertCollectionToDictCollectionKey = key2 Else convertCollectionToDictCollectionKey = key2
End Function

' used by Collection compatibility: converts a stored String key back to Collection Variant (String or Empty)
Private Function convertDictCollectionToCollectionKey(key As String) As Variant
    Dim keyType As String, keyVal As String, retVal As Variant
    retVal = key
    If Len(key) >= 6 Then
        keyType = Mid(key, 2, 3): keyVal = Mid(key, 6, Len(key) - 6)
        Select Case keyType
            Case "EMP": retVal = ""
        End Select
    End If
    convertDictCollectionToCollectionKey = retVal
End Function
' ======================== END: EMULATION HELPER FUNCTION ===========================================================================================



' ======================== START: ADVANCED FUNCTIONALITY ============================================================================================
' Moves an item from an index to another index, chainable
Public Function Move(FromIndex As Long, ToIndex As Long) As DictCollection
    Dim fromIndex2 As Long, toIndex2 As Long
    ' adjust offset
    fromIndex2 = FromIndex - mExternalIndexOffset
    toIndex2 = ToIndex - mExternalIndexOffset
    If toIndex2 < 0 Or fromIndex2 < 0 Or fromIndex2 >= mItemCount Then
        If mThrowErrors Then
            ' Emulate Collection/Array behavior
            Err.Raise 9, "DictCollection", "Subscript out of range"
        End If
        Exit Function
    End If
    If toIndex2 = fromIndex2 Then Exit Function ' exit if item does not need to be moved
    Dim KeyIndex As Long, tempItem As Variant, moveUpwards As Boolean, lower As Long, upper As Long
    ' find out direction of copy
    moveUpwards = fromIndex2 < toIndex2
    If toIndex2 >= mItemCount Then increaseItemArrayTo toIndex2
    ' remember item to move and its key index
    If IsObject(allItems(fromIndex2)) Then Set tempItem = allItems(fromIndex2) Else tempItem = allItems(fromIndex2)
    KeyIndex = allItemKeyIndexes(fromIndex2)
    If moveUpwards Then
        ' item will be moved up, all other items are copied down
        lower = fromIndex2 + 1: upper = toIndex2
        copyItemsDown lower, upper, 1, True
        If IsObject(tempItem) Then Set allItems(upper) = tempItem Else allItems(upper) = tempItem
        allItemKeyIndexes(upper) = KeyIndex ' change item-to-key pointer
    Else
        ' item will be moved down, all other items are copied up
        lower = toIndex2: upper = fromIndex2 - 1
        copyItemsUp lower, upper, 1, True
        If IsObject(tempItem) Then Set allItems(lower) = tempItem Else allItems(lower) = tempItem
        allItemKeyIndexes(lower) = KeyIndex ' change item-to-key pointer
    End If
    If (KeyIndex > -1) And (KeyIndex < mKeyCount) Then allKeyItemIndexes(KeyIndex) = toIndex2   ' change key-to-item pointer
    Set Move = Me
End Function

' adds (inserts) a new DictCollection at given index or key and returns it; will copy parent DictCollection Settings
' if IndexOrKey is Missing or "", a new DictCollection will be added after last item
Public Function AddDC(Optional key As Variant, Optional AtIndex As Long = -1) As DictCollection
    Dim newDictCollection As DictCollection
    Set newDictCollection = New DictCollection: CopyAllSettingsExceptCollectionType Me, newDictCollection
    Me.Insert newDictCollection, AtIndex, key
    Set AddDC = newDictCollection
End Function

' retrieves a DictCollection from a given index or key; adds a DictCollection if the item is nonexisting or not a DictCollection
Public Function subDc(IndexOrKey As Variant) As DictCollection
    Dim newOrExistingDictCollection As DictCollection, result As EvaluateIndexOrKeyResult, action As Integer ' 0 = found, 1 = found/overwrite with new, 2 = not found/add
    If IsMissing(IndexOrKey) And Not mEmulateScriptingDictionary Then
        ' Emulate Collection behavior
        If mThrowErrors Then Err.Raise 5, "DictCollection", "Invalid procedure call or argument"
        Exit Function ' return with return value = Nothing
    Else
        If mItemCount = -1 Then initializeItemArray
        If mKeyCount > 1 And mLazySorting And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
        If mEmulateScriptingDictionary Then
            result = evaluateIndexOrKey(convertScriptingDictionaryToDictCollectionKey(IndexOrKey))
        ElseIf mEmulateCollection Then
            result = evaluateIndexOrKey(convertCollectionToDictCollectionKey(IndexOrKey))
        Else
            result = evaluateIndexOrKey(IndexOrKey)
        End If
        If Not result.WasIndex And Not result.WasKey Then
            If mThrowErrors Then Err.Raise 13, "DictCollection", "Keys of type " & TypeName(IndexOrKey) & " are not supported!"
            Exit Function ' return with return value = Nothing
        End If
        If result.Index < 0 And result.WasIndex Then
            If mThrowErrors Then Err.Raise 9, "Subscript out of range"
            Exit Function
        ElseIf result.Index > (mItemCount - 1) And result.WasIndex Then
            If mThrowErrors Then Err.Raise 9, "Subscript out of range"
            action = 2 ' not found/add
        ElseIf result.WasFound Then
            ' key or index was found
            If IsObject(allItems(result.Index)) Then
                If InStr(TypeName(allItems(result.Index)), "DictCollection") > 0 Then
                    action = 0 ' found
                Else
                    action = 1 ' found/overwrite with new
                End If
            Else
                action = 1 ' found/overwrite with new
            End If
        Else
            ' key was not found
            action = 2 ' not found/add
        End If
    End If
    Select Case action
        Case 0:
            ' found, get existing DictCollection
            Set newOrExistingDictCollection = allItems(result.Index)
        Case 1:
            ' found, overwrite with new DictCollection
            Set newOrExistingDictCollection = New DictCollection: CopyAllSettingsExceptCollectionType Me, newOrExistingDictCollection
            Set allItems(result.Index) = newOrExistingDictCollection
        Case 2:
            ' not found, add new DictCollection
            Set newOrExistingDictCollection = New DictCollection: CopyAllSettingsExceptCollectionType Me, newOrExistingDictCollection
            If result.WasIndex Then
                ' add at given index without key
                internalAddItem newOrExistingDictCollection, , result.Index
            Else
                ' add with key
                Dim result2 As EvaluateKeyResult
                If mEmulateScriptingDictionary Then
                    result2.key = convertScriptingDictionaryToDictCollectionKey(IndexOrKey)
                    result2.WasInvalid = False: result2.WasString = True ' explicitly set to String to bypass error handling below
                ElseIf mEmulateCollection Then
                    result2 = evaluateKey(convertCollectionToDictCollectionKey(IndexOrKey))
                Else
                    result2 = evaluateKey(IndexOrKey)
                End If
                If result2.WasInvalid Then
                    If mThrowErrors Then Err.Raise 13, "DictCollection", "Keys of type " & TypeName(IndexOrKey) & " are not supported"
                    Exit Function
                End If
                internalAddItem newOrExistingDictCollection, result2.key
                If mEmulateScriptingDictionary Then storeScriptingDictionaryKeyInformation IndexOrKey, result2.key
            End If
    End Select
    Set subDc = newOrExistingDictCollection
End Function
' sets the keys of all items that have no key to "_POS" where POS = item index + 1, e.g. key for item at 12 = "_12"
' if "_POS" is already taken, a counter C ranging from 1 to 1000 is added as "_POS_C" until the new key is not taken
' does not work with .EmulateDictionary=True
Public Sub EnsureAllItemsHaveKeys()
    Dim i As Long, keyToTest As String, keyCounter As Long, keySubCounter As Long, result As FindKeyResult
    ' sort if necessary
    If mLazySorting And mKeyCount > 1 And (Not mKeysAreSortedAndUnique) Then sortKeysAndRemoveDuplicates: mKeysAreSortedAndUnique = True
    ' initialize key array if necessary
    If mKeyCount < 0 Then initializeKeyArray
    For i = 0 To (mItemCount - 1)
        keyCounter = i + 1
        If allItemKeyIndexes(i) = -1 Then
            ' test key
            keyToTest = "_" & CStr(keyCounter)
            result = internalFindKeyIndex(keyToTest, mCompareMode)
            keySubCounter = 1
            While result.Exact And keySubCounter <= 1000
                ' test key with subnumber
                keyToTest = "_" & keyCounter & "_" & CStr(keySubCounter)
                result = internalFindKeyIndex(keyToTest, mCompareMode)
                keySubCounter = keySubCounter + 1
            Wend
            If (Not result.Exact) And (keySubCounter <= 1000) Then
                ' add key using eager sorting
                insertKey keyToTest, result.Index, i
                allItemKeyIndexes(i) = result.Index
            Else
                If mThrowErrors Then Err.Raise 5, "DictCollection", "Could not assign key to item at " & i & ". Tried keys from '" & "_" & i & "' to '" & keyToTest & "'.": Exit Sub
            End If
        End If
    Next
End Sub

' Copies items of this DictCollection to TargetCollection from given targetIndex by overwriting existing items.
' targetIndex = -1 will copy all items to the end of TargetCollection, preserving all existing items
Public Function CopyItems(TargetCollection As DictCollection, Optional targetIndex As Long = -1) As DictCollection
    internalCopyItems Me, TargetCollection, targetIndex
    Set CopyItems = Me
End Function

' Copies items of this DictCollection to TargetCollection from given targetIndex by overwriting existing items.
' targetIndex = -1 will copy all items to the end of TargetCollection, preserving all existing items
Public Function CopyItemsToNew() As DictCollection
    Dim newDc As DictCollection
    internalCopyItems Me, newDc, -1
    Set CopyItemsToNew = newDc
End Function

' will copy all items and keys to given TargetCollection (or a new DictCollection) that will be returned
' if no SourceCollection is passed, this DictCollection will be taken as SourceCollection
' if no TargetCollection is passed, the SourceCollections settings will be copied to the TargetCollection
' targetIndex = -1 will copy all items to the end of TargetCollection, preserving all existing items
Private Function internalCopyItems(Optional SourceCollection As DictCollection, Optional TargetCollection As DictCollection, Optional targetIndex As Long = -1)
    Dim sourceIndex As Long, targetIndex2 As Long, tempKey As Variant, tempItem As Variant
    Dim throwErrorsWasTrue As Boolean, upper As Long, lower As Long, useAdd As Boolean
    If SourceCollection Is Nothing Then Set SourceCollection = Me
    If TargetCollection Is Nothing Then
        Set TargetCollection = New DictCollection: CopyAllSettingsExceptCollectionType SourceCollection, TargetCollection
    End If
    throwErrorsWasTrue = TargetCollection.ThrowErrors
    ' ensure correct indexes
    If SourceCollection.ZeroBasedIndex Then
        lower = 0: upper = SourceCollection.Count - 1
    Else
        lower = 1: upper = SourceCollection.Count
    End If
    
    ' TargetCollection.ZeroBasedIndex=False -> (TargetCollection.ZeroBasedIndex * 1) = 0
    ' TargetCollection.ZeroBasedIndex=True -> (TargetCollection.ZeroBasedIndex * -1) = -1
    useAdd = (targetIndex = -1) Or (targetIndex = (TargetCollection.Count - (TargetCollection.ZeroBasedIndex * -1)) + 1)
    
    TargetCollection.ThrowErrors = False
    If useAdd Then
        ' add all items
        For sourceIndex = lower To upper
            ' avoid evaluation of default property by using set
            UtilAssignFromTo SourceCollection.ItemAt(sourceIndex), tempItem
            If SourceCollection.ItemHasKey(sourceIndex) Then
                UtilAssignFromTo SourceCollection.KeyOfItemAt(sourceIndex), tempKey
                TargetCollection.Add tempKey, tempItem ' add with key
            Else
                TargetCollection.Add , tempItem ' add without key
            End If
        Next
    Else
        ' overwrite/set all items
        targetIndex2 = targetIndex
        For sourceIndex = lower To upper
            ' avoid evaluation of default property by using set
            If IsObject(SourceCollection.ItemAt(sourceIndex)) Then
                Set TargetCollection.ItemAt(targetIndex2) = SourceCollection.ItemAt(sourceIndex)
            Else
                TargetCollection.ItemAt(targetIndex2) = SourceCollection.ItemAt(sourceIndex)
            End If
            If SourceCollection.ItemHasKey(sourceIndex) Then
                If IsObject(SourceCollection.KeyOfItemAt(sourceIndex)) Then
                    Set TargetCollection.KeyOfItemAt(targetIndex2) = SourceCollection.KeyOfItemAt(sourceIndex)
                Else
                    TargetCollection.KeyOfItemAt(targetIndex2) = SourceCollection.KeyOfItemAt(sourceIndex)
                End If
            Else
                TargetCollection.KeyOfItemAt(targetIndex2) = Empty
            End If
            targetIndex2 = targetIndex2 + 1
        Next
    End If
    If throwErrorsWasTrue Then TargetCollection.ThrowErrors = True
End Function

' runs over all subcollections recoursively and merges their items into this DictCollection
' keys are concatenated into the root collection using the delimiter "." in the format "ROOT_KEY.SUB_KEY.SUB_SUB_KEY"
' items without keys will have an underscore as key followed their position in the item array: "_POS" ("_1" to "_N")
' if "_POS" is already taken, a counter C ranging from 1 to 1000 is added as "_POS_C" until the new key is not taken
' will disable any emulations in this DictCollection
' chainable (returns this DictCollection)
Public Function Flatten() As DictCollection
    Dim Options As WalkDictCollectionTreeOptions: Options.Flatten = True: Options.EnsureKeys = True
    walkDictCollectionTree Options
    Set Flatten = Me
End Function

' runs over all subcollections recoursively and merges their items into a new DictCollection
' keys are concatenated into the root collection using the delimiter "." in the format "ROOT_KEY.SUB_KEY.SUB_SUB_KEY"
' items without keys will have an underscore as key followed their position in the item array: "_POS" ("_1" to "_N")
' if "_POS" is already taken, a counter C ranging from 1 to 1000 is added as "_POS_C" until the new key is not taken
' will disable any emulations in this DictCollection
' will copy all other settings to new DictCollection
Public Function FlattenToNew() As DictCollection
    Dim Options As WalkDictCollectionTreeOptions: Options.Flatten = True: Options.EnsureKeys = True
    Dim TargetCollection As New DictCollection
    walkDictCollectionTree Options, , TargetCollection
    Set FlattenToNew = TargetCollection
End Function

' restores all flattened subcollections in this DictCollection
' chainable (returns this DictCollection)
Public Function Unflatten() As DictCollection
    Dim Options As WalkDictCollectionTreeOptions: Options.Unflatten = True
    walkDictCollectionTree Options
    Set Unflatten = Me
End Function

' restores all flattened subcollections to a new DictCollection
Public Function UnflattenToNew() As DictCollection
    Dim TargetCollection As New DictCollection
    Dim Options As WalkDictCollectionTreeOptions: Options.Unflatten = True
    walkDictCollectionTree Options, , TargetCollection
    Set UnflattenToNew = TargetCollection
End Function

' makes a deep copy of all keys/items and subcollections to new DictCollection, also copies settings
' Variant object pointers and Objects as items/keys are replicated as pointers and will still point to the same object after cloning
Public Function CloneToNew() As DictCollection
    Dim Options As WalkDictCollectionTreeOptions: Options.Clone = True
    Dim TargetCollection As New DictCollection
    walkDictCollectionTree Options, , TargetCollection ' will also copy settings
    Set CloneToNew = TargetCollection
End Function

' makes a deep copy of all keys/items and subcollections, also copies settings
' chainable (returns this DictCollection)
' Variant object pointers and Objects as items/keys are replicated as pointers and will still point to the same object after cloning
Public Function CloneTo(TargetCollection As DictCollection) As DictCollection
    Dim Options As WalkDictCollectionTreeOptions: Options.Clone = True
    TargetCollection.RemoveAll
    walkDictCollectionTree Options, , TargetCollection ' will also copy settings
    Set CloneTo = Me
End Function

' makes a deep copy of all keys/items and subcollections but keeps existing keys/items/settings
' chainable (returns this DictCollection)
' Variant object pointers and Objects as items/keys are replicated as pointers and will still point to the same object after cloning
Public Function CloneToPreserve(TargetCollection As DictCollection) As DictCollection
    Dim Options As WalkDictCollectionTreeOptions: Options.Clone = True
    Dim tmpDictCollection As New DictCollection
    Me.CopyAllSettingsExceptCollectionType TargetCollection, tmpDictCollection
    walkDictCollectionTree Options, , TargetCollection ' will temporarily copy settings
    Me.CopyAllSettingsExceptCollectionType tmpDictCollection, TargetCollection
    Set CloneToPreserve = Me
End Function

' returns a DictCollection containing these keys:
' "NonStringConvertableItems" - the number of the items in the subtree that are not convertable to String
' "NonStringConvertableKeys" - the number of the keys in the subtree that are not convertable to String
' "SubCollections" - the number of DictCollections in the subtree (without counting items below circular references)
' "SubItems" - the number of items in the subtree (without counting items below circular references)
' "CircularReferences" - the number of DictCollections in the subtree that contain DictCollections of upper levels of the same tree
' "TreeDepth" - the number subtree levels counted upwards from 0 (None)
Public Function AnalyzeDictCollectionTree(Optional Recursive As Boolean = True) As DictCollection
    Dim Options As WalkDictCollectionTreeOptions: Options.Analyze = True
    Set AnalyzeDictCollectionTree = walkDictCollectionTree(Options)
End Function


' Unflattening algorithm:
' 1) if toTargetCollection, make copy of all items to new SourceCollection
' 2) run over all items of SourceCollection
' 2.1) if item has nested key with "." delimiter then
'       - see if parent key (the part before the first ".") with subcollection already exists in Source- or TargetCollection
'       - if yes, add item with ots subkey to this subcollection and remove the item from Source- or TargetCollecton
'       - if no, create subcollection, add the current item to it with its subkey and replace the current item with the new subcollection
' 2.2) if item has an array key in the form of "_POS" or "_POS_COUNTER"
'       - check if this is the only array key for this position (array key has the highest COUNTER)
'       - if yes, remember item and its target index
' 2.3) if item has nested key and parent key is array key, do 2.1) and 2.2)
' 3) run over all subcollections and call Unflattening (recoursively)
' 4) run over all array items (2.2 or 2.3), remove them from Source- or TargetCollection and insert them back at their correct target index
Friend Function walkDictCollectionTree(Options As WalkDictCollectionTreeOptions, Optional CollectionsInPath As Variant, Optional TargetCollection As DictCollection) As DictCollection
    Dim i As Long, j As Long, tname As String, retVal As DictCollection, maxTreeDepth As Long, SourceCollection As DictCollection, SourceCollectionSettings As DictCollection
    Dim isItemStringConvertable As Boolean, isKeyStringConvertable As Boolean, parentKey As String, currentItem As Variant, currentKey As Variant, currentItemHasKey As Boolean
    Dim toTargetCollection As Boolean, tempSubCollections As Variant, tempItemIndexes As Variant, tempVar As Variant
    Dim dc2 As DictCollection, dc3 As DictCollection, otherRetVal As DictCollection, sourceCollectionHadOneBasedIndex As Boolean, sourceCollectionHadThrowErrorsTrue As Boolean
    Dim nextKeyToTest As String, numberAsText As String, asciicode As Integer, textIsNumber As Boolean, targetIndex As Long, currentIndex As Long, subCollectionExists As Boolean, subCollectionKey As String
    
    toTargetCollection = Not TargetCollection Is Nothing
    If Options.Analyze Then
        Set retVal = New DictCollection: retVal.CompareMode = vbTextCompare
        retVal.ItemOf("NonStringConvertableItems") = 0
        retVal.ItemOf("NonStringConvertableKeys") = 0
        retVal.ItemOf("SubCollections") = 0
        retVal.ItemOf("SubItems") = 0
        retVal.ItemOf("CircularReferences") = 0
        retVal.ItemOf("TreeDepth") = 0
    End If
        
    If toTargetCollection Then
        CopyAllSettingsExceptCollectionType Me, TargetCollection
    End If
    If mItemCount < 1 Then GoTo ExitFunction ' exit if nothing to do
        
    If (Options.EnsureKeys Or Options.Unflatten) And toTargetCollection Then
        Set SourceCollection = Me.CopyItemsToNew ' make safety copy, will also copy settings if called like this
    Else
        Set SourceCollection = Me
    End If
    ' temporarily switch off one based index and throw errors (will be switched back at the end of this function)
    sourceCollectionHadOneBasedIndex = Not SourceCollection.ZeroBasedIndex
    If sourceCollectionHadOneBasedIndex Then
        SourceCollection.ZeroBasedIndex = True
        If toTargetCollection Then TargetCollection.ZeroBasedIndex = True
    End If
    sourceCollectionHadThrowErrorsTrue = SourceCollection.ThrowErrors
    If sourceCollectionHadThrowErrorsTrue Then
        SourceCollection.ThrowErrors = False
        If toTargetCollection Then TargetCollection.ThrowErrors = False
    End If
    
    ' add source collection to path to be able to detect circular references in recoursive calls
    UtilAddArrayValue CollectionsInPath, SourceCollection
    If Options.Analyze Then retVal.Item("SubItems") = mItemCount
    maxTreeDepth = 0
            
    If Options.EnsureKeys Then SourceCollection.EnsureAllItemsHaveKeys
    
    ' disable any emulation
    If Options.Flatten Or Options.Unflatten Then
        If toTargetCollection Then
            If TargetCollection.EmulateDictionary Then TargetCollection.EmulateDictionary = False
            If TargetCollection.EmulateCollection Then TargetCollection.EmulateCollection = False
            ' make backup of source collection settings
            Set SourceCollectionSettings = New DictCollection: CopyAllSettingsExceptCollectionType SourceCollection, SourceCollectionSettings
        End If
        If SourceCollection.EmulateDictionary Then SourceCollection.EmulateDictionary = False
        If SourceCollection.EmulateCollection Then SourceCollection.EmulateCollection = False
    End If
    
    i = 0 ' run over all items in SourceCollection from index 0 to .Count - 1
    While i <= (SourceCollection.Count - 1)
        UtilAssignFromTo SourceCollection.ItemAt(i), currentItem
        currentItemHasKey = SourceCollection.ItemHasKey(i)
        tname = TypeName(currentItem)
        If Options.Analyze Then
            ' see if item can be converted to String or is DictCollection
            isItemStringConvertable = False
            Select Case tname
                Case "String", "Integer", "Long", "Single", "Double", "Currency", "Decimal", "Date", "Null", "Nothing", "DictCollection", "Empty": isItemStringConvertable = True
            End Select
            If Not isItemStringConvertable Then retVal.Increment "NonStringConvertableItems"
        End If
        If currentItemHasKey Then
            If SourceCollection.EmulateDictionary Then
                ' key can be object, check before assigning
                UtilAssignFromTo SourceCollection.KeyOfItemAt(i), currentKey
            Else
                ' key can be string or empty (.EmulateCollection=True)
                currentKey = SourceCollection.KeyOfItemAt(i)
            End If
            If Options.Analyze Then
                ' see if key of item can be converted to String
                isKeyStringConvertable = False
                Select Case TypeName(currentKey)
                    Case "String", "Integer", "Long", "Single", "Double", "Currency", "Decimal", "Date", "Null", "Empty": isKeyStringConvertable = True
                End Select
                If Not isKeyStringConvertable Then retVal.Increment "NonStringConvertableKeys"
            End If
        End If
        
        If tname = "DictCollection" Then
            ' if item is DictCollection, handle it recousively
            If maxTreeDepth = 0 Then maxTreeDepth = 1
            Dim index2 As Long, subTreeDepth As Long
            If Options.Analyze Then retVal.Increment "SubCollections"
            index2 = UtilFindArrayIndex(CollectionsInPath, SourceCollection.ItemAt(i))
            If index2 > -1 Then
                ' circular reference: this DictCollection already exists on upper level of nested tree
                retVal.Increment "CircularReferences"
            Else
                Set dc2 = currentItem
                If Options.Clone Then
                    ' add new empty DictCollection to TargetCollection
                    Set dc3 = New DictCollection
                    If currentItemHasKey Then
                        TargetCollection.Add currentKey, dc3 ' add with key
                    Else
                        TargetCollection.Add , dc3 ' add without key
                    End If
                ElseIf Options.Flatten And toTargetCollection Then
                    Set dc3 = New DictCollection
                End If
    ' ------------------------------------------- HERE IS THE FIRST RECOURSIVE CALL -------------------------------------------
                Set otherRetVal = dc2.walkDictCollectionTree(Options, CollectionsInPath, dc3)
                If Options.Analyze Then
                    retVal.Increment "NonStringConvertableItems", otherRetVal("NonStringConvertableItems")
                    retVal.Increment "NonStringConvertableKeys", otherRetVal("NonStringConvertableKeys")
                    retVal.Increment "SubCollections", otherRetVal("SubCollections")
                    retVal.Increment "CircularReferences", otherRetVal("CircularReferences")
                    retVal.Increment "SubItems", otherRetVal("SubItems")
                    subTreeDepth = otherRetVal("TreeDepth"): If subTreeDepth + 1 > maxTreeDepth Then maxTreeDepth = subTreeDepth + 1
                End If
                If Options.Flatten Then
                    If dc2.Count = 0 Then
                        ' empty collection, replace with value
                        If TargetCollection Is Nothing Then
                            SourceCollection.ItemAt(i) = SourceCollection.EmptyCollectionValue
                        Else
                            If currentItemHasKey Then
                                TargetCollection.Add currentKey, TargetCollection.EmptyCollectionValue
                            Else
                                TargetCollection.Add , TargetCollection.EmptyCollectionValue
                            End If
                        End If
                    Else
                        If currentItemHasKey Then parentKey = currentKey Else parentKey = "?" & i & "?"
                        If TargetCollection Is Nothing Then
                            SourceCollection.Remove i ' remove the DictCollection itself
                            For j = dc2.Count - 1 To 0 Step -1
                                SourceCollection.Insert dc2.ItemAt(j), i, parentKey & "." & dc2.KeyOfItemAt(j) ' insert to this DictCollection
                            Next
                            ' increment i by the number of items in dc2, minus 1 for the removed dc2
                             i = i + dc2.Count - 1
                        Else
                            For j = 0 To dc3.Count - 1
                                TargetCollection.Add parentKey & "." & dc3.KeyOfItemAt(j), dc3.ItemAt(j) ' add to target DictCollection
                            Next
                        End If
                    End If
                End If
            End If
        Else
            ' not DictCollection
            If (Options.Clone Or Options.Flatten Or Options.Unflatten) And toTargetCollection Then
                ' copy items
                If currentItemHasKey Then
                    TargetCollection.Add currentKey, currentItem ' add with key
                Else
                    TargetCollection.Add , currentItem ' add without key
                End If
            End If
            If Options.Unflatten And currentItemHasKey And TypeName(currentKey) = "String" Then
                ' collect item with concatenated key in subCollections and remove it in SourceCollection
                ' if item has a valid array key ("_POS" or "_POS_COUNTER") then collect it in arrayItems and its target index in arrayItemIndexes and remove it
                Dim Pos As Long, isArrayItem As Boolean, isSubCollection As Boolean, keyPartToCheckForArrayKey As String, subKey As String
                
                ' 1) check if item is subcollection
                isSubCollection = False
                Pos = InStr(1, currentKey, ".")
                If Pos > 1 Then
                    ' this is a concatenated key, get parent key and add item to temp collection
                    parentKey = Left(currentKey, Pos - 1)
                    subKey = Right(currentKey, Len(currentKey) - Pos)
                    isSubCollection = True
                Else
                    parentKey = currentKey ' check entire key
                End If
                
                ' 2) check if item is an array item (had no key before flattening)
                isArrayItem = False
                If Left(parentKey, 1) = "_" Then
                    ' get the POS part of "_POS_COUNTER" or "_POS"
                    Pos = InStr(2, parentKey, "_")
                    If Pos > 2 Then
                        numberAsText = Mid(parentKey, 2, Pos - 2)
                    Else
                        numberAsText = Mid(parentKey, 2, Len(parentKey) - 1)
                    End If
                    ' build next key by extracting numbers
                    textIsNumber = UtilStringConsistsOfNumericAsciiChars(numberAsText)
                    If textIsNumber Then
                        ' POS is a number
                        ' get target index (the index that the item had flatten operation)
                        targetIndex = CLng(numberAsText) - 1
                        ' generate next key to test
                        If Pos > 2 Then
                            ' this is a key in the form "_POS_COUNTER", increment COUNTER
                            numberAsText = Right(parentKey, Len(parentKey) - Pos)
                            ' test if COUNTER can be converted to number
                            textIsNumber = UtilStringConsistsOfNumericAsciiChars(numberAsText)
                            If textIsNumber Then nextKeyToTest = Left(parentKey, Pos) & CLng(numberAsText) + 1 ' build next key to test with format "_POS_[C+1]"
                        Else
                            ' this is a key in the form "_POS", add "_1"
                            nextKeyToTest = parentKey & "_1"
                        End If
                        ' if next key is not found in SourceCollection -> isArrayItem = True
                        ' this has the weakness that if we deliberately assign a key in the form "_P_C" before flattening,
                        ' the unflattening will detect this as an array item and thus remove the key and put the item at a wrong position P
                        If textIsNumber Then
                            ' find all keys that start with nextKeyToTest
                            tempVar = SourceCollection.FindKeysThatStartWith(nextKeyToTest)
                            isArrayItem = (UBound(tempVar) = -1) ' is array item if tempVar is empty array
                            ' if not false, continue search in TargetItems
                            If isArrayItem And toTargetCollection Then isArrayItem = Not TargetCollection.Exists(nextKeyToTest)
                        End If
                    End If
                End If
                
                '3) create new subcollection if needed, add item to it and remember array target index
                If isSubCollection Then
                    If toTargetCollection Then subCollectionExists = TargetCollection.Exists(parentKey) Else subCollectionExists = SourceCollection.Exists(parentKey)
                    If subCollectionExists Then
                        If toTargetCollection Then
                            TargetCollection.AsDC(parentKey).Add subKey, currentItem
                            TargetCollection.Remove TargetCollection.Count - 1
                        Else
                            SourceCollection.AsDC(parentKey).Add subKey, currentItem
                            SourceCollection.Remove i
                            i = i - 1 ' adjust i to reflect the removed item
                        End If
                    Else
                        ' create new DictCollection with added item
                        Set dc2 = New DictCollection
                        UtilAddArrayValue tempSubCollections, dc2 ' remember subcollection
                        dc2.Add subKey, currentItem
                        CopyAllSettingsExceptCollectionType SourceCollection, dc2
                        If sourceCollectionHadOneBasedIndex Then dc2.ZeroBasedIndex = False
                        If sourceCollectionHadThrowErrorsTrue Then dc2.ThrowErrors = True
                        If toTargetCollection Then
                            ' replace last item with new subcollection
                            Set TargetCollection.ItemAt(TargetCollection.Count - 1) = dc2
                            TargetCollection.KeyOfItemAt(TargetCollection.Count - 1) = parentKey
                        Else
                            ' replace current item with new subcollection
                            Set SourceCollection.ItemAt(i) = dc2
                            SourceCollection.KeyOfItemAt(i) = parentKey
                        End If
                    End If
                End If
                If isArrayItem Then
                    If isSubCollection Then
                        ' add subCollection if item is first item of subCollection (subCollection does not exist yet)
                        If Not subCollectionExists Then UtilAddArrayValue tempItemIndexes, Array(targetIndex, parentKey, dc2)
                    Else
                        ' add if item is no subCollection
                        UtilAddArrayValue tempItemIndexes, Array(targetIndex, parentKey, currentItem)
                    End If
                End If
            End If
        End If
        i = i + 1
    Wend
    ' remove keys of array items and move them to their original positions
    If Options.Unflatten Then
        If Not IsEmpty(tempSubCollections) Then
            For i = 0 To UBound(tempSubCollections)
                Set dc2 = tempSubCollections(i)
    ' ------------------------------------------- HERE IS THE SECOND RECOURSIVE CALL -------------------------------------------
                dc2.walkDictCollectionTree Options, CollectionsInPath
            Next
        End If
        If Not IsEmpty(tempItemIndexes) Then
            If toTargetCollection Then Set dc3 = TargetCollection Else Set dc3 = SourceCollection
            ' sort temp array by target index
            UtilSortArray tempItemIndexes, 0, UBound(tempItemIndexes)
            ' remove all array items
            For i = 0 To UBound(tempItemIndexes): dc3.Remove tempItemIndexes(i)(1): Next
            ' insert them back at target indexes
            For i = 0 To UBound(tempItemIndexes)
                UtilAssignFromTo tempItemIndexes(i)(2), currentItem
                targetIndex = tempItemIndexes(i)(0)
                dc3.Insert currentItem, targetIndex
            Next
        End If
    End If
    ' restore settings in source collection if target collection was passed
    If Options.Flatten And toTargetCollection Then CopyAllSettingsExceptCollectionType SourceCollectionSettings, SourceCollectionSettings
    If Options.Analyze Then retVal("TreeDepth") = maxTreeDepth
    UtilRemoveArrayValue CollectionsInPath, SourceCollection
            
    ' switch back one based index or throw erros settings
    If sourceCollectionHadOneBasedIndex Then
        SourceCollection.ZeroBasedIndex = False
        If toTargetCollection Then TargetCollection.ZeroBasedIndex = False
    End If
    If sourceCollectionHadThrowErrorsTrue Then
        SourceCollection.ThrowErrors = True
        If toTargetCollection Then TargetCollection.ThrowErrors = True
    End If
            
ExitFunction:
    Set walkDictCollectionTree = retVal
End Function

' will add Amount (positive or negative) to the item at index or key and return the result
' if the item does not exists and .ThrowErrors=False, Amount will be assigned to the item
Public Function Increment(IndexOrKey As Variant, Optional Amount As Variant = 1&) As Variant
    Dim result As EvaluateIndexOrKeyResult, retVal As Variant, tname As String, tname2 As String, hasDecimalPlaces As Boolean
    tname = TypeName(Amount)
    retVal = Empty
    If Not (IsNumeric(Amount) Or IsDate(Amount)) Then
        If mThrowErrors Then Err.Raise 13, "Amount of type " & TypeName(Amount) & " is not supported"
        GoTo ExitFunction
    End If
    result = internalGetItem(IndexOrKey, False)
    If Not result.WasFound Then
        ' errors will be thrown by internalGetItem
        retVal = Amount: Me.Item(IndexOrKey) = retVal
    Else
        tname2 = TypeName(allItems(result.Index))
        Select Case tname2
            Case "Integer", "Long", "Single", "Double", "Currency", "Decimal", "Date":
                retVal = allItems(result.Index) + Amount ' use implicit type conversion
            Case Else:
                retVal = Amount
        End Select
        allItems(result.Index) = retVal
    End If
ExitFunction:
    Increment = retVal
End Function

' ======================== END: ADVANCED FUNCTIONALITY ============================================================================================


' ======================== START: PUBLIC HELPER FUNCTIONS =========================================================================================

' UNTESTED
' assigns FromValue to ToValue by using a 'Set ... = ...' expression if FromValue is an Object and normal '... = ...' assignment if not
Public Function UtilAssignFromTo(FromVariable As Variant, ToVariable As Variant)
    If IsObject(FromVariable) Then Set ToVariable = FromVariable Else ToVariable = FromVariable
End Function

' UNTESTED
' takes an array of strings that are split by an asterisk character and returns True if it matches TextToCheck
' examples: StringMatchesWildCard("my_pony_is_green", Array("","pony","green"), vbCompareText) -> True (starts with or contains 'pony' and ends with 'green')
'           StringMatchesWildCard("my_pony_is_green", Array("my","blue",""), vbCompareText) -> False (starts with 'my' but does not contain or end with 'blue')
Public Function UtilStringMatchesWildcard(TextToCheck As String, WildcardSplit As Variant, Optional CompareMode As VbCompareMethod = vbBinaryCompare) As Boolean
    Dim textToFind As String, match As Boolean, i As Long, nextStartPos As Long, foundPos As Long
    If Not IsArray(WildcardSplit) Then
        match = False
    ElseIf UBound(WildcardSplit) = -1 Then
        match = False
    ElseIf Len(TextToCheck) = 0 Then
        match = False
    ElseIf UBound(WildcardSplit) = 0 Then
        ' no asterisks used, text has to match exactly
        match = (StrComp(TextToCheck, WildcardSplit(0), CompareMode) = 0)
    Else
        match = True: nextStartPos = 1
        For i = 0 To UBound(WildcardSplit)
            textToFind = WildcardSplit(i)
            If Len(textToFind) > 0 Then
                If Len(TextToCheck) < ((nextStartPos - 1) + Len(textToFind)) Then
                    match = False: Exit For  ' filename not long enough for comparison
                Else
                    Select Case i
                        Case 0:
                            ' string has to start with textToFind
                            If StrComp(Left(TextToCheck, Len(textToFind)), textToFind, CompareMode) <> 0 Then match = False: Exit For
                            nextStartPos = Len(textToFind) + 1
                        Case UBound(WildcardSplit):
                            ' string has to end with textToFind
                            If StrComp(Right(TextToCheck, Len(textToFind)), textToFind, CompareMode) <> 0 Then match = False: Exit For
                        Case Else:
                            foundPos = InStr(nextStartPos, TextToCheck, textToFind, CompareMode) ' 0 = not found
                            If foundPos >= nextStartPos Then
                                nextStartPos = foundPos + Len(textToFind)
                            Else
                                match = False: Exit For
                            End If
                    End Select
                End If
            End If
        Next
    End If
    UtilStringMatchesWildcard = match
End Function

' UNTESTED
' returns True if Text starts with SearchText
' returns False if Text does not start with SearchText or if SearchText is Empty/ZeroLengthString
Public Function UtilStringStartsWith(Text As String, SearchText As String, Optional CompareMode As VbCompareMethod = vbBinaryCompare) As Boolean
    Dim textLength As Long, searchTextLength As Long
    searchTextLength = Len(SearchText)
    If searchTextLength = 0 Then
        UtilStringStartsWith = False
    Else
        textLength = Len(Text)
        If textLength >= searchTextLength Then
            ' Mid() is faster than Left() because it is just a pointer
            UtilStringStartsWith = StrComp(Mid(Text, 1, searchTextLength), SearchText, CompareMode) = 0
        Else
            UtilStringStartsWith = False
        End If
    End If
End Function

' UNTESTED
' returns the position of the first occurence of any character in chars or -1 if none of the characters in chars was found
Public Function UtilStringFindChars(InputString As String, StartPos As Long, Chars As String) As Long
    Dim Pos As Long: UtilStringFindChars = -1
    If Len(Chars) = 0 Or Len(InputString) = 0 Then Exit Function
    For Pos = StartPos To Len(InputString)
        If InStr(1, Chars, Mid$(InputString, Pos, 1)) > 0 Then UtilStringFindChars = Pos: Exit Function
    Next
End Function

' UNTESTED
' returns true if the given string contains only (and at least one) numeric ascii characters, otherwise false
Public Function UtilStringConsistsOfNumericAsciiChars(Text As String) As Boolean
    Dim i As Long, retVal As Boolean, textLength As Long, asciicode As Integer
    textLength = Len(Text)
    If textLength > 0 Then
        retVal = True
        For i = 1 To textLength
            asciicode = AscW(Mid(Text, i, 1))
            If asciicode < 48 Or asciicode > 57 Then retVal = False: Exit For
        Next
    Else
        retVal = False
    End If
    UtilStringConsistsOfNumericAsciiChars = retVal
End Function

' UNTESTED
' returns a concatenated String that contains Text multiple times
' returns Text if NumberOfTimes is 1 and "" if NumberOfTimes is less 1
Public Function UtilStringRepeat(Text As String, NumberOfTimes As Long) As String
    Dim i As Long, textLength As Long
    textLength = Len(Text)
    If textLength > 0 Then
        If NumberOfTimes < 1 Then
            UtilStringRepeat = ""
        Else
            UtilStringRepeat = Space(textLength * NumberOfTimes)
            For i = 1 To NumberOfTimes
                Mid(UtilStringRepeat, (textLength * (i - 1)), textLength) = Text
            Next
        End If
    Else
        UtilStringRepeat = ""
    End If
End Function

' UNTESTED
' Source: http://www.cpearson.com/excel/vbaarrays.htm (modified)
' Returns the number of dimensions of an array.
' Nested 1-dimensional array = -1, unallocated dynamic array = 0
' 1-dimensional array = 1, 2-dimensional array = 2
Public Function UtilArrayDimensions(Arr As Variant) As Integer
    Dim tempVar As Variant, dimensionIndex As Integer
    On Error Resume Next
    Do
        dimensionIndex = dimensionIndex + 1: tempVar = UBound(Arr, dimensionIndex)
    Loop Until Err.Number <> 0
    Err.Clear: dimensionIndex = dimensionIndex - 1
    If dimensionIndex = 1 Then
        ' check if array is nested array and if yes, set dimensions to -1
        If UBound(Arr) > -1 Then
            If IsArray(Arr(0)) Then dimensionIndex = -1
        End If
    End If
    UtilArrayDimensions = dimensionIndex
End Function

' UNTESTED
' converts any input array to a two-dimensional array of given size by keeping all values and filling new ones with FillValue
' one-dimensional input array becomes first column

' Converts any input array to a two-dimensional array of given size by keeping all values.
' A one-dimensional input array becomes the first column. All other values will be filled by FillValue.
' Nested one-dimensional arrays will be transformed into rows (nested arrays) and columns (nested values).
' Does not handle input array with more that 2 dimensions. Does not throw errors
Public Function UtilArrayConvertTo2Dim(Arr As Variant, RowCount As Long, ColumnCount As Long, FillValue As Variant) As Variant
    Dim lower1 As Long, upper1 As Long, lower2 As Long, upper2 As Long, arr2 As Variant
    Dim dimensions As Long, i As Long, j As Long
    dimensions = UtilArrayDimensions(Arr)
    If dimensions > 0 Then
        lower1 = LBound(Arr, 1): upper1 = UBound(Arr, 1)
        If upper1 = -1 Then dimensions = 0
    End If
    ' ensure valid dimensions: at least 1 row and 1 column
    If RowCount < 1 Then RowCount = 1
    If ColumnCount < 1 Then ColumnCount = 1
    ReDim arr2(RowCount - 1, ColumnCount - 1)
    Select Case dimensions
        Case -1:
            ' 1-dimensional nested array
            For i = 0 To RowCount - 1
                If IsArray(Arr(lower1 + i)) Then
                    ' nested item is array, copy all values into row
                    lower2 = LBound(Arr(lower1 + i)): upper2 = UBound(Arr(lower1 + i))
                    For j = 0 To ColumnCount - 1
                        If lower2 + j <= upper2 Then arr2(i, j) = Arr(lower1 + i)(lower2 + j) Else arr2(i, j) = FillValue
                    Next
                ElseIf InStr(1, TypeName(Arr(lower1 + i)), "object") > 0 Then
                    ' nested item is object, set first column to object
                    For j = 0 To ColumnCount - 1
                        If j = 0 Then Set arr2(i, j) = Arr(lower1 + i) Else arr2(i, j) = FillValue
                    Next
                Else
                    ' nested item is something else, set first column = item
                    For j = 0 To ColumnCount - 1
                        If j = 0 Then arr2(i, j) = Arr(lower1 + i) Else arr2(i, j) = FillValue
                    Next
                End If
            Next
        Case 1:
            ' 1-dimensional array, copy values into first column
            For i = 0 To RowCount - 1
                For j = 0 To ColumnCount - 1
                    If (j = 0) And (lower1 + i <= upper1) Then
                        arr2(i, j) = Arr(lower1 + i)
                    Else
                        arr2(i, j) = FillValue
                    End If
                Next
            Next
        Case 2:
            ' two-dimensional array, expand or crop
            lower2 = LBound(Arr, 2): upper2 = UBound(Arr, 2)
            For i = 0 To RowCount - 1
                For j = 0 To ColumnCount - 1
                    If (lower1 + i <= upper1) And (lower2 + j <= upper2) Then
                        arr2(i, j) = Arr(lower1 + i, lower2 + j)
                    Else
                        arr2(i, j) = FillValue
                    End If
                Next
            Next
        Case Else:
            ' empty array or array with more than 2 dimensions, fill with FillValue
            For i = 0 To RowCount - 1
                For j = 0 To ColumnCount - 1
                    arr2(i, j) = FillValue
                Next
            Next
    End Select
    UtilArrayConvertTo2Dim = arr2
End Function

' UNTESTED
' converts (flattens) any input array to a one-dimensional array by keeping all values
Public Function UtilArrayConvertTo1Dim(Arr As Variant) As Variant
    Dim lower1 As Long, upper1 As Long, lower2 As Long, upper2 As Long, arr2 As Variant
    Dim dimensions As Long, i As Long, j As Long, RowCount As Long, colCount As Long
    dimensions = UtilArrayDimensions(Arr)
    If dimensions > 0 Then
        lower1 = LBound(Arr, 1): upper1 = UBound(Arr, 1)
        If upper1 = -1 Then dimensions = 0
    End If
    Select Case dimensions
        Case -1:
            ' 1-dimensional nested array
            For i = lower1 To upper1
                If IsArray(Arr(i)) Then
                    ' nested item is array, try to copy all values
                    lower2 = LBound(Arr(i)): upper2 = UBound(Arr(i))
                    For j = lower2 To upper2
                        UtilAddArrayValue arr2, Arr(lower1 + i)(j)
                    Next
                ElseIf InStr(1, TypeName(Arr(i)), "object") > 0 Then
                    ' nested item is object, add it
                    UtilAddArrayValue arr2, Empty
                    upper2 = UBound(arr2)
                    Set arr2(upper2) = Arr(i)
                Else
                    ' nested item is something else, add it
                    UtilAddArrayValue arr2, Arr(i)
                End If
            Next
        Case 0:
            ' empty or unititialized array
            arr2 = Array()
        Case 1:
            ' 1-dimensional array, do nothing
            arr2 = Arr
        Case 2:
            ' two-dimensional array, copy values to one-dimensional array
            lower2 = LBound(Arr, 2): upper2 = UBound(Arr, 2)
            For i = lower1 To upper1
                For j = lower2 To upper2
                    UtilAddArrayValue arr2, Arr(i, j)
                Next
            Next
    End Select
    UtilArrayConvertTo1Dim = arr2
End Function

' UNTESTED
' sorts a VBA array from smallest to largest using quicksort algorithm
' if array is one-dimensional nested array, arrays are sorted by the (first) values at arr(row)(0)
' if array is two-dimensional, rows are sorted by the (first) values at arr(row,0)
' if array is uninitialized or not an array, it will be set to empty array
' will crash if array contains objects and if it contains non-comparable types where the sort value is expected
' Original Source: https://wellsr.com/vba/2018/excel/vba-quicksort-macro-to-sort-arrays-fast/
Public Sub UtilSortArray(Arr As Variant, FromIndex As Long, ToIndex As Long)
    Dim pivotValue As Variant, swappedValue As Variant, lower As Long, upper As Long, dimensions As Integer, i As Long, colUpper As Long, continueLoop As Boolean
    dimensions = UtilArrayDimensions(Arr)
    lower = FromIndex
    upper = ToIndex
    Select Case dimensions
        Case 0: Arr = Array()
        Case 1:
            ' One-Dimensional Array
            pivotValue = Arr((FromIndex + ToIndex) \ 2)
            While (lower <= upper) 'divide
                While (Arr(lower) < pivotValue And lower < ToIndex)
                   lower = lower + 1
                Wend
                While (pivotValue < Arr(upper) And upper > FromIndex)
                   upper = upper - 1
                Wend
                If (lower <= upper) Then
                    swappedValue = Arr(lower): Arr(lower) = Arr(upper): Arr(upper) = swappedValue
                    lower = lower + 1: upper = upper - 1
                End If
            Wend
            If (FromIndex < upper) Then UtilSortArray Arr, FromIndex, upper 'conquer
            If (lower < ToIndex) Then UtilSortArray Arr, lower, ToIndex 'conquer
        Case -1:
            ' One-Dimensional Nested Array
            pivotValue = Arr((FromIndex + ToIndex) \ 2)(0)
            While (lower <= upper) 'divide
                While (Arr(lower)(0) < pivotValue And lower < ToIndex)
                   lower = lower + 1
                Wend
                While (pivotValue < Arr(upper)(0) And upper > FromIndex)
                   upper = upper - 1
                Wend
                If (lower <= upper) Then
                    swappedValue = Arr(lower): Arr(lower) = Arr(upper): Arr(upper) = swappedValue
                    lower = lower + 1: upper = upper - 1
                End If
            Wend
            If (FromIndex < upper) Then UtilSortArray Arr, FromIndex, upper 'conquer
            If (lower < ToIndex) Then UtilSortArray Arr, lower, ToIndex 'conquer
        Case 2:
            ' Two-Dimensional Array
            colUpper = UBound(Arr, 2)
            ReDim swappedValue(0 To colUpper)
            pivotValue = Arr((FromIndex + ToIndex) \ 2, 0)
            While (lower <= upper) 'divide
                While (Arr(lower, 0) < pivotValue And lower < ToIndex)
                   lower = lower + 1
                Wend
                While (pivotValue < Arr(upper, 0) And upper > FromIndex)
                   upper = upper - 1
                Wend
                If (lower <= upper) Then
                    For i = 0 To colUpper: swappedValue(i) = Arr(lower, i): Next
                    For i = 0 To colUpper: Arr(lower, i) = Arr(upper, i): Next
                    For i = 0 To colUpper: Arr(upper, i) = swappedValue(i): Next
                    lower = lower + 1: upper = upper - 1
                End If
            Wend
            If (FromIndex < upper) Then UtilSortArray Arr, FromIndex, upper 'conquer
            If (lower < ToIndex) Then UtilSortArray Arr, lower, ToIndex 'conquer
    End Select
End Sub

' UNTESTED
' sorts a VBA array from smallest to largest using StrComp with quicksort algorithm
' if array is one-dimensional nested array, arrays are sorted by the (first) values at arr(row)(0)
' if array is two-dimensional, rows are sorted by the (first) values at arr(row,0)
' if array is uninitialized or not an array, it will be set to empty array
' will crash if array contains objects and if it contains non-string-convertable types where the sort value is expected
' Original Source: https://wellsr.com/vba/2018/excel/vba-quicksort-macro-to-sort-arrays-fast/
Public Sub UtilSortStringArray(Arr As Variant, FromIndex As Long, ToIndex As Long, CompareMode As VbCompareMethod)
    Dim pivotValue As Variant, swappedValue As Variant, lower As Long, upper As Long, dimensions As Integer, i As Long, colUpper As Long
    dimensions = UtilArrayDimensions(Arr)
    lower = FromIndex
    upper = ToIndex
    Select Case dimensions
        Case 0: Arr = Array()
        Case 1:
            ' One-Dimensional Array
            pivotValue = Arr((FromIndex + ToIndex) \ 2)
            While (lower <= upper) 'divide
                While (StrComp(Arr(lower), pivotValue, CompareMode) = -1 And lower < ToIndex)
                   lower = lower + 1
                Wend
                While (StrComp(Arr(upper), pivotValue, CompareMode) = 1 And upper > FromIndex)
                   upper = upper - 1
                Wend
                If (lower <= upper) Then
                    swappedValue = Arr(lower): Arr(lower) = Arr(upper): Arr(upper) = swappedValue
                    lower = lower + 1: upper = upper - 1
                End If
            Wend
            If (FromIndex < upper) Then UtilSortStringArray Arr, FromIndex, upper, CompareMode 'conquer
            If (lower < ToIndex) Then UtilSortStringArray Arr, lower, ToIndex, CompareMode 'conquer
        Case -1:
            ' One-Dimensional Nested Array
            pivotValue = Arr((FromIndex + ToIndex) \ 2)(0)
            While (lower <= upper) 'divide
                While (StrComp(Arr(lower)(0), pivotValue, CompareMode) = -1 And lower < ToIndex)
                   lower = lower + 1
                Wend
                While (StrComp(Arr(upper)(0), pivotValue, CompareMode) = 1 And upper > FromIndex)
                   upper = upper - 1
                Wend
                If (lower <= upper) Then
                    swappedValue = Arr(lower): Arr(lower) = Arr(upper): Arr(upper) = swappedValue
                    lower = lower + 1: upper = upper - 1
                End If
            Wend
            If (FromIndex < upper) Then UtilSortArray Arr, FromIndex, upper 'conquer
            If (lower < ToIndex) Then UtilSortArray Arr, lower, ToIndex 'conquer
        Case 2:
            ' Two-Dimensional Array
            colUpper = UBound(Arr, 2)
            ReDim swappedValue(0 To colUpper)
            pivotValue = Arr((FromIndex + ToIndex) \ 2, 0)
            While (lower <= upper) 'divide
                While (StrComp(Arr(lower, 0), pivotValue, CompareMode) = -1 And lower < ToIndex)
                   lower = lower + 1
                Wend
                While (StrComp(Arr(upper, 0), pivotValue, CompareMode) = 1 And upper > FromIndex)
                   upper = upper - 1
                Wend
                If (lower <= upper) Then
                    For i = 0 To colUpper: swappedValue(i) = Arr(lower, i): Next
                    For i = 0 To colUpper: Arr(lower, i) = Arr(upper, i): Next
                    For i = 0 To colUpper: Arr(upper, i) = swappedValue(i): Next
                    lower = lower + 1: upper = upper - 1
                End If
            Wend
            If (FromIndex < upper) Then UtilSortStringArray Arr, FromIndex, upper, CompareMode 'conquer
            If (lower < ToIndex) Then UtilSortStringArray Arr, lower, ToIndex, CompareMode 'conquer
    End Select
End Sub

' UNTESTED
' randomizes the positions of the values in an array without changing them
' if array is one-dimensional nested array, contained arrays are shuffled
' if array is two-dimensional, contained rows are shuffled
' if array is uninitialized or not an array, it will be set to empty array
' will crash if array contains objects
' original source: http://www.cpearson.com/excel/ShuffleArray.aspx
Public Sub UtilShuffleArray(Arr As Variant)
    Dim oldIndex As Long, newIndex As Long, swappedValue As Variant, dimensions As Integer, lower As Long, upper As Long, i As Long, colUpper As Long
    dimensions = UtilArrayDimensions(Arr)
    Select Case dimensions
        Case 0: ' uninitialized array
            Arr = Array(): Exit Sub
        Case 1, -1: ' one-dimensional array or one-dimensional nested array
            lower = LBound(Arr): upper = UBound(Arr)
        Case 2: ' two-dimensional array
            lower = LBound(Arr, 1): upper = UBound(Arr, 1)
            colUpper = UBound(Arr, 2)
            ReDim swappedValue(0 To colUpper)
    End Select
    Randomize
    For oldIndex = lower To upper
        newIndex = CLng(((upper - oldIndex) * Rnd) + oldIndex) ' move value somewhere upwards in array
        If oldIndex <> newIndex Then
            ' swap values
            Select Case dimensions
                Case 1, -1: ' one-dimensional array or one-dimensional nested array
                    swappedValue = Arr(oldIndex): Arr(oldIndex) = Arr(newIndex): Arr(newIndex) = swappedValue
                Case 2: ' two-dimensional array
                    For i = 0 To colUpper: swappedValue(i) = Arr(oldIndex, i): Next
                    For i = 0 To colUpper: Arr(oldIndex, i) = Arr(newIndex, i): Next
                    For i = 0 To colUpper: Arr(newIndex, i) = swappedValue(i): Next
                Case 3:
            End Select
        End If
    Next
End Sub

' adds a value to an array
Public Sub UtilAddArrayValue(Arr As Variant, val As Variant)
    If InStr(TypeName(Arr), "()") < 1 Then
        ' if arr is not array, initialize array with val as value
        Arr = Array(val)
    Else
        ' increment array size by 1
        ReDim Preserve Arr(LBound(Arr) To UBound(Arr) + 1)
        ' set object reference or value
        If IsObject(val) Then Set Arr(UBound(Arr)) = val Else Arr(UBound(Arr)) = val
    End If
End Sub

' removes all occurrences of a value or object reference from an array
Public Sub UtilRemoveArrayValue(Arr As Variant, val As Variant)
    Dim i As Long, newArr As Variant, newArrIndex As Long
    If InStr(TypeName(Arr), "()") < 1 Then Arr = Array(): Exit Sub
    ReDim newArr(LBound(Arr) To UBound(Arr)): newArrIndex = 0
    For i = LBound(Arr) To UBound(Arr)
        ' copy object reference or value if not equal
        If IsObject(Arr(i)) And IsObject(val) Then
            If Not (Arr(i) Is val) Then
                Set newArr(newArrIndex) = Arr(i): newArrIndex = newArrIndex + 1
            Else
                If UBound(newArr) > LBound(newArr) Then ReDim Preserve newArr(LBound(newArr) To UBound(newArr) - 1) Else newArr = Array()
            End If
        Else
            If Arr(i) <> val Then
                newArr(newArrIndex) = Arr(i): newArrIndex = newArrIndex + 1
            Else
                If UBound(newArr) > LBound(newArr) Then ReDim Preserve newArr(LBound(newArr) To UBound(newArr) - 1) Else newArr = Array()
            End If
        End If
    Next
    Arr = newArr
End Sub

' removes an array value from the array at a given index
Public Sub UtilRemoveArrayValueByIndex(Arr As Variant, Index As Long)
    Dim i As Long
    If InStr(TypeName(Arr), "()") < 1 Or (Index = 0 And (UBound(Arr) = LBound(Arr))) Then Arr = Array(): Exit Sub
    If Index < LBound(Arr) Or Index > UBound(Arr) Then Exit Sub
    If Index < UBound(Arr) Then
        ' copy all values above index one down
        For i = Index + 1 To UBound(Arr)
            If IsObject(Arr(i)) Then Set Arr(i - 1) = Arr(i) Else Arr(i - 1) = Arr(i)
        Next
    End If
    ' decrease array length by 1
    ReDim Preserve Arr(LBound(Arr) To UBound(Arr) - 1)
End Sub

' finds first occurrence of a value or object refernce within an array or returns -1 if not found
Public Function UtilFindArrayIndex(Arr As Variant, val As Variant) As Long
    Dim i As Long
    If InStr(TypeName(Arr), "()") < 1 Then Arr = Array(): UtilFindArrayIndex = -1: Exit Function
    For i = LBound(Arr) To UBound(Arr)
        ' copy object reference or value if not equal
        If IsObject(Arr(i)) And IsObject(val) Then
            If (Arr(i) Is val) Then UtilFindArrayIndex = i: Exit Function
        Else
            If Arr(i) = val Then UtilFindArrayIndex = i: Exit Function
        End If
    Next
    UtilFindArrayIndex = -1
End Function

Public Function UtilGetMissingValue(Optional DoNotPassAnythingHere) As Variant: UtilGetMissingValue = DoNotPassAnythingHere: End Function
' ======================== END: PUBLIC HELPER FUNCTIONS ===========================================================================================



' ======================== START: DEMO AND SELFTEST ===============================================================================================
' performs demo of basic functions
Public Sub DemoBasicFunctionality()
    Dim dc As DictCollection: Set dc = New DictCollection
    Debug.Print "----------- Start DictCollection Demo of Basic Functionality -------------------"
    Debug.Print "  dc.Add ""key1"", 123.45": dc.Add "key1", 123.45
    Debug.Print "     dc.Item(0) should be 123.45 -> " & dc.Item(0)
    Debug.Print "  dc.Add, ""A""": dc.Add , "A"
    Debug.Print "     dc.Item(1) should be A -> " & dc.Item(1)
    Debug.Print "     dc.Count should be 2 -> " & dc.Count
    Debug.Print "     dc.KeyCount should be 1 -> " & dc.KeyCount
    Debug.Print "     dc.Item(""key1"") should be 123.45 -> " & dc.Item("key1")
    Debug.Print "     dc.Item(1) should be A -> " & dc.Item(1)
    Debug.Print "  dc.Item(0) = 100": dc.Item(0) = 100
    Debug.Print "     dc.Item(0) should be 100 -> " & dc.Item(0)
    Debug.Print "  dc.Item(""key1"") = 200": dc.Item("key1") = 200
    Debug.Print "     dc.Item(0) should be 200 -> " & dc.Item(0)
    Debug.Print "  dc.Item(""key2"") = ""B""": dc.Item("key2") = "B"
    Debug.Print "     dc.Item(2) should be B -> " & dc.Item(2)
    Debug.Print "  dc.Add ""key3"", New DictCollection": dc.Add "key3", New DictCollection
    Debug.Print "     TypeName(dc.Item(3)) should be DictCollection -> " & TypeName(dc.Item(3))
    Debug.Print "  dc.Item(""key3"").Add ""key4"", 456.78""": dc.Item("key3").Add "key4", 456.78
    Debug.Print "     dc.Item(3).Item(0) should be 456.78 -> " & dc.Item(3).Item(0)
    Debug.Print "  dc.Item(""key3"").Add, ""C""": dc.Item("key3").Add , "C"
    Debug.Print "     dc.Item(3).Item(1) should be C -> " & dc.Item(3).Item(1)
    Debug.Print "     dc.Item(""key3"").Item(""key4"") should be 456.78 -> " & dc.Item("key3").Item("key4")
    Debug.Print "     dc.Item(""key3"").Item(1) should be C -> " & dc.Item("key3").Item(1)
    Debug.Print "     (dc.Item(""key3"").Item(1) = dc.NonExistingValue) should be False -> " & (dc.Item("key3").Item(1) = dc.NonExistingValue)
    Debug.Print "     StrComp(dc.Item(""key3"").Item(""wrongkey""), dc.NonExistingValue) should be 0 (equal) -> " & StrComp(dc("key3")("wrongkey"), dc.NonExistingValue)
    Debug.Print "  dc.Remove(""key1"")": dc.Remove ("key1")
    Debug.Print "     dc.Item(0) should be A -> " & dc.Item(0)
    Debug.Print "  dc.Remove(0)": dc.Remove (0)
    Debug.Print "     dc.Item(0) should be B -> " & dc.Item(0)
    Debug.Print "     dc.IndexOfKey(""key3"") should be 1 -> " & dc.IndexOfKey("key3")
    Debug.Print "     dc.KeyOfItemAt(0) should be key2 -> " & dc.KeyOfItemAt(0)
    Debug.Print "  dc.Key(""key2"") = ""key4""": dc.key("key2") = "key4"
    Debug.Print "     dc.KeyOfItemAt(0) should be key4 -> " & dc.KeyOfItemAt(0)
    Debug.Print "     dc.Exists(""key3"") should be True -> " & dc.Exists("key3")
    Debug.Print "     dc.Count should be 2 -> " & dc.Count
    Debug.Print "  dc.Key(""key4"") = ""key3""": dc.key("key4") = "key3"
    Debug.Print "     dc.KeyOfItemAt(0) should be key3 -> " & dc.KeyOfItemAt(0)
    Debug.Print "     dc.Exists(""key4"") should be False -> " & dc.Exists("key4")
    Debug.Print "     dc.Count should be 1 -> " & dc.Count
    Debug.Print "     Join(dc.Keys,"","") should be key3 -> " & Join(dc.Keys, ",")
    Debug.Print "     Join(dc.Items,"","") should be B -> " & Join(dc.Items, ",")
    Debug.Print "  dc.Insert ""D"", 0, ""key5""": dc.Insert "D", 0, "key5"
    Debug.Print "     dc.Item(0) should be D -> " & dc.Item(0)
    Debug.Print "     dc.Item(1) should be B -> " & dc.Item(1)
    Debug.Print "  dc.Insert ""E"", 3, """"": dc.Insert "E", 3, ""
    Debug.Print "     Join(dc.Items,"","") should be D,B,,E -> " & Join(dc.Items, ",")
    Debug.Print "     IsEmpty(dc.Item(2)) should be True -> " & IsEmpty(dc.Item(2))
    Debug.Print "     Join(dc.Keys,"","") should be key5,key3,, -> " & Join(dc.Keys, ",")
    Debug.Print "     Join(dc.Keys(False),"","") should be key5,key3 -> " & Join(dc.Keys(False), ",")
    Debug.Print "     Join(dc.SortedKeys,"","") should be key3,key5 -> " & Join(dc.SortedKeys, ",")
    Debug.Print "  dc.Move 0, 2": dc.Move 0, 2
    Debug.Print "     Join(dc.Items,"","") should be B,,D,E -> " & Join(dc.Items, ",")
    Debug.Print "  dc.RemoveAll": dc.RemoveAll
    Debug.Print "     dc.Count should be 0 -> " & dc.Count
    Debug.Print "----------- End DictCollection Demo of Basic Functionality ---------------------"
End Sub

' performs all available tests
' IMPORTANT: If you experience unstable, nondeterministic test results, try adding a refined messsage to the currentTest string
'            Example: Change ("-> " & currentTest) to ("-> " & currentTest & ": a special condition must be met")
'            Sometimes these string operations cause the disappearance of nondeterministic results (observed when testing for correctness of thrown errors)
Public Function SelfTest(Optional DebugPrint As Boolean = True)
    Dim retVal As Boolean: retVal = True
    retVal = retVal And UBound(TestFunctionality(DebugPrint)) > -1
    retVal = retVal And UBound(TestCompatibility(DebugPrint)) > -1
    SelfTest = retVal
End Function

' adds a value to errors and prints it to Immediate window if DebugPrint=true
Private Sub logError(DebugPrint As Boolean, errors As Variant, val As String)
    If InStr(TypeName(errors), "()") < 1 Then errors = Array()
    ' increment size by one
    ReDim Preserve errors(LBound(errors) To UBound(errors) + 1)
    ' set value and print if necessary
    errors(UBound(errors)) = val
    If DebugPrint Then Debug.Print "  " & val
End Sub

' combines errors with allErrors and prints errors to Immediate window if if DebugPrint=true
Private Function combineAndPrintTestErrors(testName As String, errors As Variant, allErrors As Variant, Optional DebugPrint As Boolean = True)
    Dim i As Long
    Debug.Print IIf(UBound(errors) > -1, "  FAIL: ", "  OK: ") & testName
    For i = 0 To UBound(errors)
        UtilAddArrayValue allErrors, errors(i): If DebugPrint Then Debug.Print "  " & errors(i)
    Next
End Function
' ======================== END: DEMO AND SELFTEST =================================================================================================



' ======================== START: FUNCTIONALITY TESTS =============================================================================================

' performs selftest and returns all errors as variant array
Public Function TestFunctionality(Optional DebugPrint As Boolean = True) As Variant
    Dim allErrors As Variant: allErrors = Array():
    
    If DebugPrint Then Debug.Print "----------- Start DictCollection Functionality Test ----------------------------"
    
    combineAndPrintTestErrors "Basic Functionality", testBasicFunctionality(False), allErrors, DebugPrint
    combineAndPrintTestErrors "Advanced Item And Key Access", testAdvancedItemAndKeyAccess(False), allErrors, DebugPrint
    combineAndPrintTestErrors "Advanced Functionality 1", testAdvancedFunctionality1(False), allErrors, DebugPrint
    combineAndPrintTestErrors "Advanced Functionality 2", testAdvancedFunctionality2(False), allErrors, DebugPrint
    combineAndPrintTestErrors "Utility Functions", testUtilFunctions(False), allErrors, DebugPrint
    combineAndPrintTestErrors "Bugfixes and Workarounds", testFixes(False), allErrors, DebugPrint
    
    If DebugPrint Then Debug.Print "----------- End DictCollection Functionality Test ------------------------------"
    TestFunctionality = allErrors
End Function

' Tests basic functionality (DebugPrint=True -> print errors immediately)
' dc.CollectionType, dc.CompareMode
' dc.Add(val), dc.Add(val, key)
' dc.Item(key), dc.Item(index), dc(), dc(key), dc(index)
' dc.Remove(key), dc.Remove(index), dc.RemoveAll
Private Function testBasicFunctionality(Optional DebugPrint As Boolean = False) As Variant
    On Error GoTo Fail: Err.Clear
    Dim errors As Variant, dc1 As DictCollection, dc2 As DictCollection, currentTest As String, i As Long, val1 As Variant, val2 As Variant, key1 As Variant, key2 As Variant
    errors = Array()
    
    Set dc1 = New DictCollection ' initialize class
currentTest = "[IOP-1] Initialized object properties - New DictCollection must have .Count=0": If dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    currentTest = "[IOP-2] Initialized object properties - New DictCollection must have .CollectionType=1 (empty array)": If dc1.CollectionType <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
    currentTest = "[IOP-3] Initialized object properties - Initial CompareMode must be binary": If dc1.CompareMode <> VbCompareMethod.vbBinaryCompare Then logError DebugPrint, errors, ("-> " & currentTest)
    currentTest = "[IOP-4] Initialized object properties - Throw Errors must be fale by default": If dc1.ThrowErrors <> False Then logError DebugPrint, errors, ("-> " & currentTest)
    currentTest = "[SAG-1] Simple add and get - Adding a single item": dc1.Add "key1", "item1"
    currentTest = "[SAG-2] Simple add and get - Retrieving single item": If dc1("key1") <> "item1" Then logError DebugPrint, errors, ("-> " & currentTest)
    currentTest = "[SAG-3] Simple add and get - Retrieving key for single item": If dc1.KeyOfItemAt(0) <> "key1" Then logError DebugPrint, errors, ("-> " & currentTest)
    currentTest = "[SAG-4] Simple add and get - Retrieving index for key": If dc1.IndexOfKey("key1") <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    
    Set dc1 = New DictCollection: dc1.Add , 100 ' initialize class and add value
currentTest = "[AOP-1] Array object properties - DictCollection with one item must have .Count=1": If dc1.Count <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
    currentTest = "[AOP-2] Array object properties - DictCollection with one item and no keys must have .CollectionType=2 (filled array)": If dc1.CollectionType <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    
    Set dc1 = New DictCollection: dc1.Add "keyA", 100 ' initialize class and add value with key
currentTest = "[COP-1] Collection object properties - DictCollection with one item and one key must have .CollectionType=4 (filled key-value-store)": If dc1.CollectionType <> 4 Then logError DebugPrint, errors, ("-> " & currentTest)
    Set dc1 = New DictCollection: dc1.Add "keyA", 100: dc1.Add , 200 ' initialize class and add value with key and value without key
currentTest = "[COP-2] Collection object properties - DictCollection with one item and one key must have .CollectionType=5 (filled key-value-store with items having no key)": If dc1.CollectionType <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    
    Set dc1 = New DictCollection ' initialize class
currentTest = "[DV-1] Default values - Default NonExistingValue must be '[NONEXISTING]'": If dc1.NonExistingValue <> NONEXISTING_VALUE_DEFAULT Then logError DebugPrint, errors, ("-> " & currentTest)
    currentTest = "[DV-2] Default values - Default EmptyCollectionValue must be '[EMPTY]'": If dc1.EmptyCollectionValue <> EMPTY_COLLECTION_VALUE_DEFAULT Then logError DebugPrint, errors, ("-> " & currentTest)
    currentTest = "[DV-3] Default values - Empty DictCollection must have '[EMPTY]' as default value": If dc1 <> dc1.EmptyCollectionValue Then logError DebugPrint, errors, ("-> " & currentTest)
    currentTest = "[DV-4] Default values - Nonexisting item must have '[NONEXISTING]' as default value": If dc1(0) <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    currentTest = "[DV-5] Default values - Nonexisting item must have '[NONEXISTING]' as default value": If dc1("nonexisting") <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    
currentTest = "[MDV-1] Modified default values for nonexisting items and empty collections"
    Set dc1 = New DictCollection: dc1.NonExistingValue = 1: dc1.EmptyCollectionValue = 2 ' initialize class and modify default values
    If dc1.NonExistingValue <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.EmptyCollectionValue <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1 <> 2 Or dc1 <> dc1.EmptyCollectionValue Then logError DebugPrint, errors, ("-> " & currentTest)
    If CInt(dc1(0)) <> 1 Or dc1(0) <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    If CInt(dc1(1)(2)(3)(4)) <> 1 Or dc1(1)(2)(3)(4) <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    If CInt(dc1("nonexisting")) <> 1 Or dc1("nonexisting") <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    If CInt(dc1("a")("b")("c")("d")) <> 1 Or dc1("a")("b")("c")("d") <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    If CInt(dc1("")) <> 1 Or dc1("") <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[MDV-2] Modified default values in nested DictCollections"
    dc1.Add "keyA", New DictCollection: dc1(0).Add "keyB", New DictCollection
    'If CInt(dc1(0)) <> 2 Or dc1(0) <> dc1.EmptyCollectionValue Then logError DebugPrint, errors, ("-> " & currentTest)

    If CInt(dc1(0)) <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    If CInt(dc1("keyA")) <> 2 Or dc1("keyA") <> dc1.EmptyCollectionValue Then logError DebugPrint, errors, ("-> " & currentTest)
    If CInt(dc1(0)(0)) <> 2 Or dc1(0)(0) <> dc1.EmptyCollectionValue Then logError DebugPrint, errors, ("-> " & currentTest)
    If CInt(dc1("keyA")("keyB")) <> 2 Or dc1("keyA")("keyB") <> dc1.EmptyCollectionValue Then logError DebugPrint, errors, ("-> " & currentTest)
    If CInt(dc1(0)(0)(1)) <> 1 Or dc1(0)(0)(1) <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    If CInt(dc1("keyA")("keyB")("nonexisting")) <> 1 Or dc1("keyA")("keyB")("nonexisting") <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    
    Set dc1 = New DictCollection
    Dim testarray(1) As Variant: testarray(0) = "TEST": testarray(1) = 222
    Set dc2 = New DictCollection: dc2.Add , "COL": dc2.Add , 333
    
currentTest = "[AIWAK-1] Adding items with ascending keys"
    dc1.RemoveAll: dc1.ThrowErrors = False: dc1.CompareMode = VbCompareMethod.vbBinaryCompare
    dc1.Add "keyA", "itemA"
    dc1.Add "keyB", 111
    dc1.Add "keyC", testarray
    dc1.Add "keyD", dc2
currentTest = "[AIWAK-2] Retrieving items with ascending keys by index":
    If dc1.Item(0) <> "itemA" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(1) <> 111 Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(2)(0) <> testarray(0) Or dc1(2)(1) <> testarray(1) Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(3)(1) <> dc2(1) Or dc1(3)(2) <> dc2(2) Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[AIWAK-3] Retrieving items with ascending keys by key"
    If dc1.Item("keyA") <> "itemA" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("keyB") <> 111 Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("keyC")(0) <> testarray(0) Or dc1("keyC")(1) <> testarray(1) Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("keyD")(0) <> dc2(0) Or dc1("keyD")(1) <> dc2(1) Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[AIWAK-4] Counting items with ascending keys": If dc1.Count <> 4 Then logError DebugPrint, errors, ("-> " & currentTest)
    
currentTest = "[AIWDK-1] Adding items with descending keys"
    dc1.RemoveAll: dc1.ThrowErrors = False: dc1.CompareMode = VbCompareMethod.vbBinaryCompare
    dc1.Add "keyC", "itemC"
    dc1.Add "keyB", "itemB"
    dc1.Add "keyA", "itemA"
currentTest = "[AIWDK-2] Retrieving items with descending keys by index"
    If dc1(0) <> "itemC" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(1) <> "itemB" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(2) <> "itemA" Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[AIWDK-3] Retrieving items with descending keys by key"
    If dc1("keyA") <> "itemA" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("keyB") <> "itemB" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("keyC") <> "itemC" Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[AIWDK-4] Counting items with descending keys": If dc1.Count <> 3 Then logError DebugPrint, errors, ("-> " & currentTest)
    
currentTest = "[AIWWK-1] Adding items with and without keys"
    dc1.RemoveAll: dc1.ThrowErrors = False: dc1.CompareMode = VbCompareMethod.vbBinaryCompare
    dc1.Add , "FirstArrayItem"
    dc1.Add "key1", "FirstKeyItem"
    dc1.Add "", "SecondArrayItem"
    dc1.Add "key2", "SecondKeyItem"
currentTest = "[AIWWK-2] Retrieving items items with and without keys by index"
    If dc1(0) <> "FirstArrayItem" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(1) <> "FirstKeyItem" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(2) <> "SecondArrayItem" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(3) <> "SecondKeyItem" Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[AIWWK-3] Retrieving items items with and without keys by index"
    If dc1("key1") <> "FirstKeyItem" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("key2") <> "SecondKeyItem" Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[AIWWK-4] Counting items with and without keys": If dc1.Count <> 4 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[AIWWK-5] Counting keys of items with and without keys": If dc1.KeyCount <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    
currentTest = "[AII-1] Accessing invalid items"
    dc1.RemoveAll: dc1.ThrowErrors = False: dc1.CompareMode = VbCompareMethod.vbBinaryCompare
    dc1.Add "keyC", "itemC"
    dc1.Add "keyB", "itemB"
    dc1.Add "keyA", "itemA"
currentTest = "[AII-2] Retrieving items with invalid index"
    If dc1(-1) <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(3) <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[AII-3] Retrieving items with invalid index"
    If dc1("") <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("_") <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.NonExistingValue = "": val1 = dc1("keyc")
    If val1 <> "" Then logError DebugPrint, errors, ("-> " & currentTest)
    'If dc1("keyc") <> "" Then logError DebugPrint, errors, ("-> " & currentTest) ' <- This does not work. Left and right side of <> operator have to be either both properties or both values
    
currentTest = "[TCM-1] Testing binary CompareMode - Adding items"
    dc1.RemoveAll: dc1.ThrowErrors = False: dc1.CompareMode = VbCompareMethod.vbBinaryCompare
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
    dc1.Add "a‡·‚„‰Â", "LOWER_ASCII_A_CHARS" 'lower case variants of the ASCII character a
    dc1.Add "A¿¡¬√ƒ≈", "UPPER_ASCII_A_CHARS" 'upper case variants of the ASCII character a
    dc1.Add ChrW(257) & ChrW(259) & ChrW(261), "LOWER_UNICODE_A_CHARS"  'lower case variants of the UNICODE character a
    dc1.Add ChrW(256) & ChrW(258) & ChrW(260), "UPPER_UNICODE_A_CHARS"  'upper case variants of the UNICODE character a
    If dc1.Count <> 16 Or dc1.KeyCount <> 16 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[TCM-3] Retrieving items with case sensitive key comparison"
    If dc1("keya") <> "itema" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("keyA") <> "itemA" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("keyB") <> "itemB" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("KeYB") <> dc1.NonExistingValue Or dc1("KEYB") <> dc1.NonExistingValue Or dc1("keyb") <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("KEYC") <> "ITEMC" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("KEyC") <> dc1.NonExistingValue Or dc1("Keyc") <> dc1.NonExistingValue Or dc1("keyc") <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(Chr(128)) <> "EURO128" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(Chr(164)) <> "EURO164" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(ChrW(8352)) <> "EURO8352" Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Add "Ä", "EURO" ' adding euro again with symbol
    If dc1.Count <> 16 Or dc1.KeyCount <> 16 Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(Chr(128)) <> "EURO" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("Ä") <> "EURO" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(ChrW(183)) <> "BULLET1" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(ChrW(8226)) <> "BULLET2" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(ChrW(8729)) <> "BULLET3" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(ChrW(9679)) <> "BULLET4" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(ChrW(9702)) <> "BULLET5" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("a‡·‚„‰Â") <> "LOWER_ASCII_A_CHARS" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("A¿¡¬√ƒ≈") <> "UPPER_ASCII_A_CHARS" Then logError DebugPrint, errors, ("-> " & currentTest) ' upper case keys will be sorted before lower case keys
    If dc1(ChrW(257) & ChrW(259) & ChrW(261)) <> "LOWER_UNICODE_A_CHARS" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(ChrW(256) & ChrW(258) & ChrW(260)) <> "UPPER_UNICODE_A_CHARS" Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[TCM-4] Change of CompareMode must leave keys intact"
    dc1.CompareMode = VbCompareMethod.vbTextCompare ' case insensitive, converts unicode characters
    ' in VBScript, some Unicode Code Points are sorted differently than in VBA:
    ' | Index | VBA Key                         | VBScript Key
    ' +-------+---------------------------------+--------------------------------
    ' | 0 - 7 | ... same order ...              | ... same order ...
    ' |  008  | "aaa" (LOWER_UNICODE_A_CHARS)   | "a‡·‚„‰Â" (LOWER_ASCII_A_CHARS)
    ' |  009  | "AAA" (UPPER_UNICODE_A_CHARS)   | "A¿¡¬√ƒ≈" (UPPER_ASCII_A_CHARS)
    ' |  010  | "A¿¡¬√ƒ≈" (UPPER_ASCII_A_CHARS) | "AAA" (UPPER_UNICODE_A_CHARS)
    ' |  011  | "a‡·‚„‰Â" (LOWER_ASCII_A_CHARS) | "aaa" (LOWER_UNICODE_A_CHARS)
    ' |  012  | "keyA"                          | "keyA"
    ' |  013  | "keya"                          | "keya"
    ' |  014  | "keyB"                          | "keyB"
    ' |  015  | "keyC"                          | "keyC"
    ' Also, in VBScript some Code Points are not equal when comparing them in case insensitive "Text Mode"
    ' VBA:       StrComp("A¿¡¬√ƒ≈","a‡·‚„‰Â",1) ' -> 0 = equal
    ' VBScript:  StrComp("A¿¡¬√ƒ≈","a‡·‚„‰Â",1) ' -> 1 = not equal
    If dc1("keya") <> "itemA" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("keyA") <> "itemA" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("A¿¡¬√ƒ≈") <> "UPPER_ASCII_A_CHARS" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("a‡·‚„‰Â") <> "UPPER_ASCII_A_CHARS" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(ChrW(257) & ChrW(259) & ChrW(261)) <> "LOWER_UNICODE_A_CHARS" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(ChrW(256) & ChrW(258) & ChrW(260)) <> "LOWER_UNICODE_A_CHARS" Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.CompareMode = VbCompareMethod.vbBinaryCompare ' case sensitive again
    If dc1("keya") <> "itema" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("keyA") <> "itemA" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("a‡·‚„‰Â") <> "LOWER_ASCII_A_CHARS" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("A¿¡¬√ƒ≈") <> "UPPER_ASCII_A_CHARS" Then logError DebugPrint, errors, ("-> " & currentTest) ' upper case keys will be sorted before lower case keys
    If dc1(ChrW(257) & ChrW(259) & ChrW(261)) <> "LOWER_UNICODE_A_CHARS" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(ChrW(256) & ChrW(258) & ChrW(260)) <> "UPPER_UNICODE_A_CHARS" Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[TCM-5] Change of CompareMode with empty key array must work"
    dc1.RemoveAll
    dc1.CompareMode = VbCompareMethod.vbBinaryCompare: If dc1.CompareMode <> VbCompareMethod.vbBinaryCompare Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.CompareMode = VbCompareMethod.vbTextCompare: If dc1.CompareMode <> VbCompareMethod.vbTextCompare Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.CompareMode = VbCompareMethod.vbDatabaseCompare: If dc1.CompareMode <> VbCompareMethod.vbDatabaseCompare Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[TCM-6] Testing TextCompare Mode - Adding items with case insensitive keys"
    dc1.RemoveAll
    dc1.CompareMode = VbCompareMethod.vbTextCompare
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
    dc1.Add "a‡·‚„‰Â", "LOWER_ASCII_A_CHARS" 'lower case variants of the ASCII character a
    dc1.Add "A¿¡¬√ƒ≈", "UPPER_ASCII_A_CHARS" 'overwrites the item above
    dc1.Add ChrW(257) & ChrW(259) & ChrW(261), "LOWER_UNICODE_A_CHARS"  'lower case variants of the UNICODE character a
    dc1.Add ChrW(256) & ChrW(258) & ChrW(260), "UPPER_UNICODE_A_CHARS"  'overwrites the item above
    currentTest = "[TCM-6] Retrieving items with case insensitive key comparison"
    If dc1("keya") <> "ITEMA" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("KEYA") <> "ITEMA" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(0) <> "ITEMA" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("KeYb") <> "itemb" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(1) <> "itemb" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("keyc") <> "ItemC" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(2) <> "ItemC" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(Chr(128)) <> "EURO128" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(Chr(164)) <> "EURO164" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(ChrW(8352)) <> "EURO8352" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("Ä") <> "EURO128" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(ChrW(183)) <> "BULLET1" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(ChrW(8226)) <> "BULLET2" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(ChrW(8729)) <> "BULLET3" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(ChrW(9679)) <> "BULLET4" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(ChrW(9702)) <> "BULLET5" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("a‡·‚„‰Â") <> "UPPER_ASCII_A_CHARS" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("A¿¡¬√ƒ≈") <> "UPPER_ASCII_A_CHARS" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(ChrW(257) & ChrW(259) & ChrW(261)) <> "UPPER_UNICODE_A_CHARS" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(ChrW(256) & ChrW(258) & ChrW(260)) <> "UPPER_UNICODE_A_CHARS" Then logError DebugPrint, errors, ("-> " & currentTest)

currentTest = "[AOA-1] Array only access"
    dc1.RemoveAll: dc1.ThrowErrors = False: dc1.CompareMode = VbCompareMethod.vbBinaryCompare
    dc1.Add , "ItemA"
    dc1.Add , 1.23
    dc1.Add , testarray
    dc1.Add , "4"
    dc1.Add , 5.67
currentTest = "[AOA-2] Testing array item count"
    If dc1.Count <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[AOA-3] Removing array items"
    dc1.Remove 1
   If dc1.Count <> 4 Then logError DebugPrint, errors, ("-> " & currentTest)
    If (dc1(0) <> "ItemA") Or (dc1(1)(0) <> "TEST") Or (dc1(2) <> "4") Or (dc1(3) <> 5.67) Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Remove 3
    If dc1.Count <> 3 Then logError DebugPrint, errors, ("-> " & currentTest)
    If (dc1(0) <> "ItemA") Or (dc1(1)(0) <> "TEST") Or (dc1(2) <> "4") Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Remove 0.9 'should remove item at index 1 because of VBA.Collection emulated behavior
    If (dc1(0) <> "ItemA") Or (dc1(1) <> "4") Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Remove Empty 'should not remove anything
    dc1.Remove "" 'should not remove anything
    dc1.Remove 0
    If dc1.Count <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(0) <> "4" Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Remove dc1.Count - 1
    If dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Remove 0 ' should not throw errors
currentTest = "[AOA-4] Adding 1000 array items"
    For i = 0 To 1000
        dc1.Add , i * 10
    Next
currentTest = "[AOA-5] Testing access of 1000 array items"
    For i = 0 To 999 Step 3
        If dc1(i) <> i * 10 Then logError DebugPrint, errors, ("-> " & currentTest)
    Next
currentTest = "[AOA-6] Getting and setting array items"
    ' removing first 10 items and changin next 10 items to "X"
    For i = 0 To 9
        dc1.Remove 0: dc1(9) = "X"
    Next
    For i = 0 To 9
        If dc1(i) <> "X" Then logError DebugPrint, errors, ("-> " & currentTest)
    Next
    
currentTest = "[MADU-1] Mixed Array and Dictionary usage"
    dc1.RemoveAll: dc1.ThrowErrors = False: dc1.CompareMode = VbCompareMethod.vbBinaryCompare
    dc1.Add , "ItemAWithoutKey"         '0 -> will be removed by index
    dc1.Add "Key1", "ItemBWithKey1"     '1 -> will be at index=0, item="ItemFWithKey1"
    dc1.Add "", testarray               '2 -> will be at index=1, item=testArray
    dc1.Add "Key2", "ItemCWithKey2"     '3 -> will be removed by index
    dc1.Add "key3", "ItemDWithKey3"     '4 -> will be at index=2, item=Empty
    dc1.Add , "ItemEWithoutKey"         '5 -> will be at index=3, item="ItemEWithoutKey"
    dc1.Add , Empty                     '6 -> will be at index=4, item="Hello"
    dc1.Add "789", testarray            '7 -> will be at index=5, item=testArray
    dc1.Add "456", "Hello"              '8 -> will be at index=6
    dc1.Add "123", "World"              '9 -> will be removed
    dc1(1) = "ItemFWithKey1"
    dc1.Remove 3
    dc1("key3") = dc1("Key2") ' should change to empty string
    dc1.Remove (dc1.Count - 1) ' remove last element
    dc1(5) = dc1("456")
    dc1.Remove 0
    If dc1.Count <> 7 Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1("") <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    If (dc1(0) <> "ItemFWithKey1") Or (dc1("Key1") <> "ItemFWithKey1") Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(1)(1) <> 222 Then logError DebugPrint, errors, ("-> " & currentTest)
    If (dc1("Key2") <> dc1.NonExistingValue) Or dc1.Exists("Key2") Then logError DebugPrint, errors, ("-> " & currentTest)
    If (dc1(2) <> dc1.NonExistingValue) Or (Not dc1.Exists("key3")) Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(3) <> "ItemEWithoutKey" Then logError DebugPrint, errors, ("-> " & currentTest)
    If (dc1(4) <> dc1("456")) Or (dc1("456") <> "Hello") Then logError DebugPrint, errors, ("-> " & currentTest)
    If (dc1(5)(0) <> "TEST") Or (dc1("789")(1) <> 222) Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(6) <> "Hello" Then logError DebugPrint, errors, ("-> " & currentTest)
    
    ' remove 3 items at the beginning -> itemCount=4, keyCount = 2
    dc1.Remove 0: dc1.Remove 0: dc1.Remove 0:
    If (dc1.Count <> 4) Or (dc1.KeyCount <> 2) Then logError DebugPrint, errors, ("-> " & currentTest)
    
    ' remove 2 items at the end -> itemCount=2, keyCount=0
    dc1.Remove (dc1.Count - 1): dc1.Remove (dc1.Count - 1)
    If (dc1.Count <> 2) Or (dc1.KeyCount <> 0) Then logError DebugPrint, errors, ("-> " & currentTest)
    
    ' remove remaining 2 items
    dc1.Remove (dc1.Count - 1): dc1.Remove 0
    If (dc1.Count <> 0) Or (dc1.KeyCount <> 0) Then logError DebugPrint, errors, ("-> " & currentTest)
    
currentTest = "[MODK-1] Manipulation of Dictionary keys - Replacing existing keys with existing keys"
    dc1.RemoveAll: dc1.ThrowErrors = False: dc1.CompareMode = VbCompareMethod.vbBinaryCompare
    dc1.Add "Key1", "Item1"
    dc1.Add "Key2", "Item2"
    dc1.Add "Key3", "Item3"
    dc1.Add "Key4", "Item4"
    dc1.Add "Key5", "Item5"
    dc1.key("Key1") = "Key2" ' -> "Item2" should be dropped
    dc1.key("Key5") = "Key4" ' -> "Item4" should be dropped
    dc1.key("Key2") = "Key1" ' -> "Key1" should again point to "Item1"
    dc1.key("Key4") = "Key5" ' -> "Key5" should again point to "Item5"
    If (dc1.Count <> 3) Or (dc1.KeyCount <> 3) Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(0) <> "Item1" Or dc1("Key1") <> "Item1" Or dc1.KeyOfItemAt(0) <> "Key1" Or dc1.IndexOfKey("Key1") <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(1) <> "Item3" Or dc1("Key3") <> "Item3" Or dc1.KeyOfItemAt(1) <> "Key3" Or dc1.IndexOfKey("Key3") <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(2) <> "Item5" Or dc1("Key5") <> "Item5" Or dc1.KeyOfItemAt(2) <> "Key5" Or dc1.IndexOfKey("Key5") <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[MODK-2] Replacing existing and nonexisting keys with nonexisting keys"
    dc1.key("Key1") = "xyz"
    dc1.key("Key2") = "efg" ' should do nothing because "Key2" does not exist anymore
    dc1.key("Key3") = "abc"
    dc1.key("Key5") = "abcd"
    If dc1(0) <> "Item1" Or dc1("xyz") <> "Item1" Or dc1.KeyOfItemAt(0) <> "xyz" Or dc1.IndexOfKey("xyz") <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(1) <> "Item3" Or dc1("abc") <> "Item3" Or dc1.KeyOfItemAt(1) <> "abc" Or dc1.IndexOfKey("abc") <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(2) <> "Item5" Or dc1("abcd") <> "Item5" Or dc1.KeyOfItemAt(2) <> "abcd" Or dc1.IndexOfKey("abcd") <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    If (dc1.Count <> 3) Or (dc1.KeyCount <> 3) Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.key("abcd") = ""
    If (dc1.Count <> 3) Or (dc1.KeyCount <> 2) Or dc1("abcd") <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.key("abc") = ""
    dc1.key("xyz") = Empty
    If (dc1.Count <> 3) Or (dc1.KeyCount <> 0) Then logError DebugPrint, errors, ("-> " & currentTest)
    
currentTest = "[ISRII-1] Inserting, Setting and Removing Items by Index - Inserting Items by Index"
    dc1.RemoveAll: dc1.ThrowErrors = False: dc1.CompareMode = VbCompareMethod.vbBinaryCompare
    dc1.Insert "e", 0   ' -> inserted at 0
    dc1.Insert "c", 0   ' -> inserted before "e" at 0, "e" is now at 1
    dc1.Insert "d", 1   ' -> inserted before "e" at 1, "e" is not at 2
    dc1.Insert "f", 3   ' -> inserted after "e" at 3
    dc1.Insert "b", 0   ' -> inserted at 0 before "c", all others move one up
    dc1.Insert "a", 0   ' -> inserted at 0 before "b", all others move one up
    dc1.Insert "h", 7   ' -> inserted at 7, creates empty item at 6
    If dc1(0) <> "a" Or dc1(1) <> "b" Or dc1(2) <> "c" Or dc1(3) <> "d" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1(4) <> "e" Or dc1(5) <> "f" Or dc1(6) <> Empty Or dc1(7) <> "h" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.Count <> 8 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[ISRII-2] Adding and Removing Items by Index - Setting Items by Index"
    dc1.RemoveAll: dc1.ThrowErrors = False: dc1.CompareMode = VbCompareMethod.vbBinaryCompare
    dc1(3) = "d": dc1(1) = "b": dc1(0) = "a": dc1(4) = "e"
    If dc1(0) <> "a" Or dc1(1) <> "b" Or dc1(2) <> Empty Or dc1(3) <> "d" Or dc1(4) <> "e" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.Count <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[ISRII-3] Adding and Removing Items by Index - Removing Items by Index"
    dc1.Remove 4: If dc1(3) <> "d" Or dc1(4) <> dc1.NonExistingValue Or dc1.Count <> 4 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Remove 2: If dc1(1) <> "b" Or dc1(2) <> "d" Or dc1(3) <> dc1.NonExistingValue Or dc1.Count <> 3 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Remove 0: If dc1(0) <> "b" Or dc1(1) <> "d" Or dc1(2) <> dc1.NonExistingValue Or dc1.Count <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Remove 0: If dc1(0) <> "d" Or dc1(1) <> dc1.NonExistingValue Or dc1.Count <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Remove 0: If dc1(0) <> dc1.NonExistingValue Or dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Remove 0: If dc1(0) <> dc1.NonExistingValue Or dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    
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
        If dc1(3) <> 111 Then logError DebugPrint, errors, ("-> " & currentTest & " dc(3) must equal 111")
        If dc1(0) <> 222 Then logError DebugPrint, errors, ("-> " & currentTest & " dc(0) must equal 222")
        If dc1(2) <> 333 Then logError DebugPrint, errors, ("-> " & currentTest & " dc(2) must equal 333")
        If dc1(1) <> 444 Then logError DebugPrint, errors, ("-> " & currentTest & " dc(1) must equal 444")
        If dc1(4) <> 555 Then logError DebugPrint, errors, ("-> " & currentTest & " dc(4) must equal 555")
        If dc1(5) <> 666 Then logError DebugPrint, errors, ("-> " & currentTest & " dc(5) must equal 666")
        If dc1(6) <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest & " dc(6) must equal " & dc1.NonExistingValue)
        If dc1("a") <> 111 Then logError DebugPrint, errors, ("-> " & currentTest & " dc(""a"") must equal 111")
        If dc1("b") <> 222 Then logError DebugPrint, errors, ("-> " & currentTest & " dc(""b"") must equal 222")
        If dc1("c") <> 333 Then logError DebugPrint, errors, ("-> " & currentTest & " dc(""c"") must equal 333")
        If dc1("d") <> 444 Then logError DebugPrint, errors, ("-> " & currentTest & " dc(""d"") must equal 444")
        If dc1("e") <> 555 Then logError DebugPrint, errors, ("-> " & currentTest & " dc(""e"") must equal 555")
        If dc1("f") <> 666 Then logError DebugPrint, errors, ("-> " & currentTest & " dc(""f"") must equal 666")
        If dc1("x") <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest & " dc(""x"") must equal " & dc1.NonExistingValue)

        If i = 1 Then
currentTest = "[CK-1] Changing keys with eager sorting": Set dc1 = New DictCollection: dc1.LazySorting = False
        Else
currentTest = "[CK-2] Changing keys with lazy sorting": Set dc1 = New DictCollection: dc1.LazySorting = True
        End If

        dc1.RemoveAll: dc1.Add "keyA", "item1": dc1.Add "keyC", "item2": dc1.Add "keyE", "item3": dc1.Add "keyF", "item4"
        key1 = "keyA": val1 = "item1": If dc1(key1) <> val1 Then logError DebugPrint, errors, ("-> " & currentTest & " dc(""" & key1 & """) must equal """ & val1 & """")
        key1 = "keyC": val1 = "item2": If dc1(key1) <> val1 Then logError DebugPrint, errors, ("-> " & currentTest & " dc(""" & key1 & """) must equal """ & val1 & """")
        ' change keys to unused keys
        dc1.key("keyA") = "keyB": key1 = "keyB": val1 = "item1": If dc1(key1) <> val1 Then logError DebugPrint, errors, ("-> " & currentTest & " dc(""" & key1 & """) must equal """ & val1 & """ after change")
        dc1.key("keyC") = "keyD": key1 = "keyD": val1 = "item2": If dc1(key1) <> val1 Then logError DebugPrint, errors, ("-> " & currentTest & " dc(""" & key1 & """) must equal """ & val1 & """ after change")
        dc1.key("keyF") = "keyA": key1 = "keyA": val1 = "item4": If dc1(key1) <> val1 Then logError DebugPrint, errors, ("-> " & currentTest & " dc(""" & key1 & """) must equal """ & val1 & """ after change")
        dc1.key("keyE") = "keyF": key1 = "keyF": val1 = "item3": If dc1(key1) <> val1 Then logError DebugPrint, errors, ("-> " & currentTest & " dc(""" & key1 & """) must equal """ & val1 & """ after change")
        dc1.key("keyB") = "keyG": key1 = "keyG": val1 = "item1": If dc1(key1) <> val1 Then logError DebugPrint, errors, ("-> " & currentTest & " dc(""" & key1 & """) must equal """ & val1 & """ after change")
        If dc1("keyG") <> "item1" Or dc1("keyD") <> "item2" Or dc1("keyA") <> "item4" Or dc1("keyF") <> "item3" Then logError DebugPrint, errors, ("-> wrong key/item associations after change.")
        ' change keys to used keys -> drop items that used those keys before
        ' item1 will get "keyA" and item4 will be dropped, still existing after that:  "keyA"=item1, "keyD"=item2, "keyF"=item3
        dc1.key("keyG") = "keyA": key1 = "keyA": val1 = "item1": If dc1(key1) <> val1 Or dc1.Exists("keyG") Or (dc1.Count <> 3) Or (dc1.KeyCount <> 3) Then logError DebugPrint, errors, ("-> " & currentTest & " item4 should not exist after assigning its key to item1")
        ' item3 will get "keyD" and item2 will be dropped, still existing after that: "keyA"=item1, "keyD"="item3
        dc1.key("keyF") = "keyD": key1 = "keyD": val1 = "item3": If dc1(key1) <> val1 Or dc1.Exists("keyF") Or (dc1.Count <> 2) Or (dc1.KeyCount <> 2) Then logError DebugPrint, errors, ("-> " & currentTest & " item2 should not exist after assigning its key to item1")
        ' item3 will get "keyA" and item1 will be dropped, still existing after that: "keyD"="item3
        dc1.key("keyD") = "keyA": key1 = "keyA": val1 = "item3": If dc1(key1) <> val1 Or dc1.Exists("keyD") Or (dc1.Count <> 1) Or (dc1.KeyCount <> 1) Then logError DebugPrint, errors, ("-> " & currentTest & " item1 should not exist after assigning its key to item1")
        ' item3 key will be set to none, still existing after that: item3 without key
        dc1.key("keyA") = "": If dc1(0) <> "item3" Or (dc1.KeyCount <> 0) Then logError DebugPrint, errors, ("-> " & currentTest & ": all keys must be removed after dropping all keys by setting them")
    Next
    

    GoTo ExitFunction
Fail:
    logError DebugPrint, errors, ("-> " & currentTest & " -> Error " & Err.Number & " " & Err.Description)
ExitFunction:
    testBasicFunctionality = errors
End Function

' Tests advanced item access functionality (DebugPrint=True -> print errors immediately)
Private Function testAdvancedItemAndKeyAccess(Optional DebugPrint As Boolean = False) As Variant
    On Error GoTo Fail: Err.Clear
    Dim errors As Variant, dc1 As DictCollection, dc2 As DictCollection, currentTest As String, i As Long, j As Long, Items As Variant, Keys As Variant, testdata As Variant, val As Variant
    Dim intItems() As Variant, intKeys() As String
    errors = Array()
    
currentTest = "[AKVPFA-1] AddPairs must add nothing from incompatible arrays"
    Set dc1 = New DictCollection
    dc1.AddPairs Array()
    If dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection must be empty when adding empty array")
    dc1.RemoveAll: dc1.AddPairs Array(Array())
    If dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection must be empty when adding empty nested array")
    dc1.RemoveAll: dc1.AddPairs Array(Array("x"))
    If dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection must be empty when adding nested array with wrong dimensions")
    ReDim testdata(0, 0)
    dc1.RemoveAll: dc1.AddPairs testdata
    If dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection must be empty when adding 2dim. array with wrong dimensions (1)")
    ReDim testdata(1, 0)
    dc1.RemoveAll: dc1.AddPairs testdata
    If dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection must be empty when adding 2dim. array with wrong dimensions (2)")
    ReDim testdata(0, 1)
    dc1.RemoveAll: dc1.AddPairs testdata
    If dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection must be empty when adding 2dim. array with empty keys")
currentTest = "[AKVPFA-2A] AddPairs must add from 2-dimensional array to empty DictCollection"
    dc1.RemoveAll
    ReDim testdata(1, 1): testdata(0, 0) = "a": testdata(0, 1) = 1: testdata(1, 0) = "b": testdata(1, 1) = 2
    dc1.AddPairs testdata
    If dc1.Count <> 2 Or dc1.KeyCount <> 2 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection must have .Count=2 and .KeyCount=2")
    If dc1.KeyOfItemAt(0) <> "a" Or dc1.KeyOfItemAt(1) <> "b" Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection has incorrect keys")
    If dc1.ItemAt(0) <> 1 Or dc1.ItemAt(1) <> 2 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection has incorrect items")
currentTest = "[AKVPFA-2B] AddPairs must add from 2-dimensional array to filled DictCollection"
    dc1.RemoveAll: dc1.Add "a", 1 ' add existing item
    ReDim testdata(1, 1): testdata(0, 0) = "b": testdata(0, 1) = 2: testdata(1, 0) = "c": testdata(1, 1) = 3
    dc1.AddPairs testdata
    If dc1.Count <> 3 Or dc1.KeyCount <> 3 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection must have .Count=3 and .KeyCount=3")
    If dc1.KeyOfItemAt(0) <> "a" Or dc1.KeyOfItemAt(1) <> "b" Or dc1.KeyOfItemAt(2) <> "c" Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection has incorrect keys")
    If dc1.ItemAt(0) <> 1 Or dc1.ItemAt(1) <> 2 Or dc1.ItemAt(2) <> 3 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection has incorrect items")
currentTest = "[AKVPFA-2C] AddPairs must add from 2-dimensional array in value-key-order to filled DictCollection"
    dc1.RemoveAll: dc1.Add "a", 1 ' add existing item
    ReDim testdata(1, 1): testdata(0, 0) = 2: testdata(0, 1) = "b": testdata(1, 0) = 3: testdata(1, 1) = "c"
    dc1.AddPairs testdata, False
    If dc1.Count <> 3 Or dc1.KeyCount <> 3 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection must have .Count=3 and .KeyCount=3")
    If dc1.KeyOfItemAt(0) <> "a" Or dc1.KeyOfItemAt(1) <> "b" Or dc1.KeyOfItemAt(2) <> "c" Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection has incorrect keys")
    If dc1.ItemAt(0) <> 1 Or dc1.ItemAt(1) <> 2 Or dc1.ItemAt(2) <> 3 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection has incorrect items")
currentTest = "[AKVPFA-3A] AddPairs must add from nested array to empty DictCollection"
    dc1.RemoveAll
    dc1.AddPairs Array(Array("a", 1), Array("b", 2))
    If dc1.Count <> 2 Or dc1.KeyCount <> 2 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection must have .Count=2 and .KeyCount=2")
    If dc1.KeyOfItemAt(0) <> "a" Or dc1.KeyOfItemAt(1) <> "b" Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection has incorrect keys")
    If dc1.ItemAt(0) <> 1 Or dc1.ItemAt(1) <> 2 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection has incorrect items")
currentTest = "[AKVPFA-3B] AddPairs must add from nested array to filled DictCollection"
    dc1.RemoveAll: dc1.Add "a", 1 ' add existing item
    dc1.AddPairs Array(Array("b", 2), Array("c", 3))
    If dc1.Count <> 3 Or dc1.KeyCount <> 3 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection must have .Count=3 and .KeyCount=3")
    If dc1.KeyOfItemAt(0) <> "a" Or dc1.KeyOfItemAt(1) <> "b" Or dc1.KeyOfItemAt(2) <> "c" Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection has incorrect keys")
    If dc1.ItemAt(0) <> 1 Or dc1.ItemAt(1) <> 2 Or dc1.ItemAt(2) <> 3 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection has incorrect items")
currentTest = "[AKVPFA-3C] AddPairs must add from nested array in value-key-order to filled DictCollection"
    dc1.RemoveAll: dc1.Add "a", 1 ' add existing item
    dc1.AddPairs Array(Array(2, "b"), Array(3, "c")), False
    If dc1.Count <> 3 Or dc1.KeyCount <> 3 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection must have .Count=3 and .KeyCount=3")
    If dc1.KeyOfItemAt(0) <> "a" Or dc1.KeyOfItemAt(1) <> "b" Or dc1.KeyOfItemAt(2) <> "c" Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection has incorrect keys")
    If dc1.ItemAt(0) <> 1 Or dc1.ItemAt(1) <> 2 Or dc1.ItemAt(2) <> 3 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection has incorrect items")
currentTest = "[AKVPFA-4] AddPairs must throw correct errors if array has inorrect dimensions"
    Set dc1 = New DictCollection: dc1.ThrowErrors = True
    On Error Resume Next
    Err.Clear: dc1.AddPairs Array(): If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection must throw Error 13 when adding empty array")
    Err.Clear: dc1.RemoveAll: dc1.AddPairs Array(Array()): If Err.Number = 13 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection must throw Error 13 when adding empty nested array")
    Err.Clear: dc1.RemoveAll: dc1.AddPairs Array(Array("x")): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection must throw Error 9 when adding nested array with wrong dimensions")
    Err.Clear: ReDim testdata(0, 0): dc1.RemoveAll: dc1.AddPairs testdata: If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection must throw Error 9 when adding 2dim. array with wrong dimensions (1)")
    Err.Clear: ReDim testdata(1, 0): dc1.RemoveAll: dc1.AddPairs testdata: If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection must throw Error 9 when adding 2dim. array with wrong dimensions (2)")
    Err.Clear: ReDim testdata(0, 1): dc1.RemoveAll: dc1.AddPairs testdata: If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest & ": DictCollection must throw Error 13 when adding 2dim. array with empty keys")
    On Error GoTo Fail: Err.Clear
    
currentTest = "[IHC-1] ItemHasKey must return correct values"
    Set dc1 = New DictCollection: dc1.Add , "A": dc1.Add "b", "B"
    If dc1.ItemHasKey(-1) Or dc1.ItemHasKey(0) Or Not dc1.ItemHasKey(1) Or dc1.ItemHasKey(2) Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[IHC-2] ItemHasKey must throw correct errors"
    Set dc1 = New DictCollection: dc1.ThrowErrors = True: dc1.Add , "A": dc1.Add "b", "B"
    On Error Resume Next
    Err.Clear: val = dc1.ItemHasKey(-1): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.ItemHasKey(2): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear
    
currentTest = "[KAIS-1] KeyOfItemAtAsString must return correct values"
    Set dc1 = New DictCollection: dc1.EmulateDictionary = True: dc1.ThrowErrors = False
    Set dc2 = New DictCollection: dc2.DefaultValueEnabled = False
    testdata = Array(CInt(1), CLng(2), CSng(3.5), CDbl(4.5), CCur(5), CDec(6), "", "x", UtilGetMissingValue, dc2, Null, CVErr(11), Nothing) ' valid Dictionary key types
    For i = 0 To UBound(testdata): dc1.Add testdata(i), i: Next
    If dc1.KeyOfItemAtAsString(-1) <> "" Or dc1.KeyOfItemAtAsString(UBound(testdata) + 1) <> "" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.KeyOfItemAtAsString(0) <> "[NUM:1]" Or dc1.KeyOfItemAtAsString(1) <> "[NUM:2]" Or dc1.KeyOfItemAtAsString(2) <> "[NUM:" & CStr(3.5) & "]" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.KeyOfItemAtAsString(3) <> "[NUM:" & CStr(4.5) & "]" Or dc1.KeyOfItemAtAsString(4) <> "[NUM:5]" Or dc1.KeyOfItemAtAsString(5) <> "[NUM:6]" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.KeyOfItemAtAsString(6) <> "[EMP:]" Or dc1.KeyOfItemAtAsString(7) <> "x" Or dc1.KeyOfItemAtAsString(8) <> "[ERR:448]" Or dc1.KeyOfItemAtAsString(9) <> "[OBJ:0]" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.KeyOfItemAtAsString(10) <> "[NUL:]" Or dc1.KeyOfItemAtAsString(11) <> "[ERR:11]" Or dc1.KeyOfItemAtAsString(12) <> "[OBJ:1]" Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[KAIS-2] KeyOfItemAtAsString must throw correct errors"
    Set dc1 = New DictCollection: dc1.ThrowErrors = True
    On Error Resume Next
    Err.Clear: val = dc1.KeyOfItemAtAsString(-1): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.KeyOfItemAtAsString(2): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear

currentTest = "[KIOIA-1] KeyIndexOfItemAt must return -1 for nonexisting item indexes"
    Set dc1 = New DictCollection
    val = dc1.KeyIndexOfItemAt(-1): If val <> -1 Then logError DebugPrint, errors, ("-> " & currentTest)
    val = dc1.KeyIndexOfItemAt(0): If val <> -1 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[KIOIA-2] KeyIndexOfItemAt must throw correct errors for nonexisting item indexes"
    Set dc1 = New DictCollection: dc1.ThrowErrors = True: dc1.Add "a", "A"
    On Error Resume Next
    Err.Clear: val = dc1.KeyAtKeyIndex(-1): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.KeyAtKeyIndex(1): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear
currentTest = "[KIOIA-3] KeyIndexOfItemAt must return correct key indexes"
    Set dc1 = New DictCollection: dc1.Add "c", "C": dc1.Add "b", "B": dc1.Add "a", "A": dc1.Insert "D", 4
    If dc1.KeyIndexOfItemAt(0) <> 2 Or dc1.KeyIndexOfItemAt(1) <> 1 Or dc1.KeyIndexOfItemAt(2) <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.KeyIndexOfItemAt(3) <> -1 Or dc1.KeyIndexOfItemAt(4) <> -1 Then logError DebugPrint, errors, ("-> " & currentTest)

currentTest = "[KAKI-1] KeyAtKeyIndex must return Empty for nonexisting indexes"
    Set dc1 = New DictCollection
    If dc1.KeyAtKeyIndex(-1) <> Empty Or dc1.KeyAtKeyIndex(0) <> Empty Or dc1.KeyAtKeyIndex(1) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[KAKI-2] KeyAtKeyIndex must throw correct errors"
    Set dc1 = New DictCollection: dc1.ThrowErrors = True: dc1.Add "a", "A"
    On Error Resume Next
    Err.Clear: val = dc1.KeyAtKeyIndex(-1): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.KeyAtKeyIndex(1): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear
currentTest = "[KAKI-3] KeyAtKeyIndex must return correct key types"
    Set dc1 = New DictCollection: dc1.EmulateDictionary = True: dc1.ThrowErrors = False
    Set dc2 = New DictCollection: dc2.DefaultValueEnabled = False
    testdata = Array(CInt(1), CLng(2), CSng(3.5), CDbl(4.5), CCur(5), CDec(6), "", "x", UtilGetMissingValue, dc2, Null, CVErr(11), Nothing) ' valid Dictionary key types
    For i = 0 To UBound(testdata): dc1.Add testdata(i), i: Next
    For i = 0 To UBound(testdata)
        j = dc1.KeyIndexOfItemAt(i)
        UtilAssignFromTo dc1.KeyAtKeyIndex(j), val
        If TypeName(testdata(i)) <> TypeName(val) Then
            logError DebugPrint, errors, ("-> " & currentTest)
        Else
            If IsObject(testdata(i)) Then
                If Not val Is testdata(i) Then logError DebugPrint, errors, ("-> " & currentTest)
            Else
                If val <> testdata(i) Then logError DebugPrint, errors, ("-> " & currentTest)
            End If
        End If
    Next

currentTest = "[KAKIS-1] KeyAtKeyIndexAsString must return "" for nonexisting indexes"
    Set dc1 = New DictCollection
    If dc1.KeyAtKeyIndexAsString(-1) <> "" Or dc1.KeyAtKeyIndexAsString(0) <> "" Or dc1.KeyAtKeyIndexAsString(1) <> "" Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[KAKIS-2] KeyAtKeyIndexAsString must throw correct errors"
    Set dc1 = New DictCollection: dc1.ThrowErrors = True: dc1.Add "a", "A"
    On Error Resume Next
    Err.Clear: val = dc1.KeyAtKeyIndexAsString(-1): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.KeyAtKeyIndexAsString(1): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear
currentTest = "[KAKIS-3] KeyAtKeyIndexAsString must return correct key types"
    Set dc1 = New DictCollection: dc1.EmulateDictionary = True: dc1.ThrowErrors = False
    Set dc2 = New DictCollection: dc2.DefaultValueEnabled = False
    testdata = Array(CInt(1), CLng(2), CSng(3.5), CDbl(4.5), CCur(5), CDec(6), "", "x", UtilGetMissingValue, dc2, Null, CVErr(11), Nothing) ' valid Dictionary key types
    For i = 0 To UBound(testdata): dc1.Add testdata(i), i: Next
    If dc1.KeyAtKeyIndexAsString(dc1.KeyIndexOfItemAt(-1)) <> "" Or dc1.KeyAtKeyIndexAsString(dc1.KeyIndexOfItemAt(UBound(testdata) + 1)) <> "" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.KeyAtKeyIndexAsString(dc1.KeyIndexOfItemAt(0)) <> "[NUM:1]" Or dc1.KeyAtKeyIndexAsString(dc1.KeyIndexOfItemAt(1)) <> "[NUM:2]" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.KeyAtKeyIndexAsString(dc1.KeyIndexOfItemAt(2)) <> "[NUM:" & CStr(3.5) & "]" Or dc1.KeyAtKeyIndexAsString(dc1.KeyIndexOfItemAt(3)) <> "[NUM:" & CStr(4.5) & "]" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.KeyAtKeyIndexAsString(dc1.KeyIndexOfItemAt(4)) <> "[NUM:5]" Or dc1.KeyAtKeyIndexAsString(dc1.KeyIndexOfItemAt(5)) <> "[NUM:6]" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.KeyAtKeyIndexAsString(dc1.KeyIndexOfItemAt(6)) <> "[EMP:]" Or dc1.KeyAtKeyIndexAsString(dc1.KeyIndexOfItemAt(7)) <> "x" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.KeyAtKeyIndexAsString(dc1.KeyIndexOfItemAt(8)) <> "[ERR:448]" Or dc1.KeyAtKeyIndexAsString(dc1.KeyIndexOfItemAt(9)) <> "[OBJ:0]" Or dc1.KeyAtKeyIndexAsString(dc1.KeyIndexOfItemAt(10)) <> "[NUL:]" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.KeyAtKeyIndexAsString(dc1.KeyIndexOfItemAt(11)) <> "[ERR:11]" Or dc1.KeyAtKeyIndexAsString(dc1.KeyIndexOfItemAt(12)) <> "[OBJ:1]" Then logError DebugPrint, errors, ("-> " & currentTest)
    
currentTest = "[FSIC-1] Function SetItem must add/set items as expected"
    Set dc1 = New DictCollection
    dc1.SetItem "a", "x": dc1.SetItem 1, "B": dc1.SetItem "c", "C": dc1.SetItem UtilGetMissingValue, "A"  ' sets default (first) value if key is Missing
    Items = dc1.Items: If Items(0) <> "A" Or Items(1) <> "B" Or Items(2) <> "C" Then logError DebugPrint, errors, ("-> " & currentTest)
    Keys = dc1.Keys: If Keys(0) <> "a" Or Keys(1) <> Empty Or Keys(2) <> "c" Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[FSIC-2] Function SetItem must throw correct errors"
    Set dc1 = New DictCollection: dc1.ThrowErrors = True
    ' errors are inherited from internalSetItem()
    On Error Resume Next
    ' empty DictCollection
    Err.Clear: Set dc2 = dc1.SetItem(-1, 1): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: Set dc2 = dc1.SetItem(0, 1): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: Set dc2 = dc1.SetItem("a", 1): If Err.Number <> 91 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Add "x", 1 ' filled DictCollection
    Err.Clear: Set dc2 = dc1.SetItem(-1, 2): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: Set dc2 = dc1.SetItem("a", 2): If Err.Number <> 91 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: Set dc2 = dc1.SetItem(Nothing, 2): If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: Set dc2 = dc1.SetItem(CVErr(11), 2): If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: Set dc2 = dc1.SetItem("x", 2): If Err.Number <> 0 Or dc1.Item("x") <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear
currentTest = "[FSIC-3] Function SetItem must be chainable"
    Set dc1 = New DictCollection: dc1.SetItem("a", "x").SetItem(1, "B").SetItem("c", "C").SetItem UtilGetMissingValue, "A" ' sets default (first) value if key is Missing
    Items = dc1.Items: If Items(0) <> "A" Or Items(1) <> "B" Or Items(2) <> "C" Then logError DebugPrint, errors, ("-> " & currentTest)
    Keys = dc1.Keys: If Keys(0) <> "a" Or Keys(1) <> Empty Or Keys(2) <> "c" Then logError DebugPrint, errors, ("-> " & currentTest)


currentTest = "[FAC-1] Function Add must be chainable"
    Set dc1 = New DictCollection: dc1.Add("a", "A").Add(, "B").Add("c", "C").Add
    Items = dc1.Items: If Items(0) <> "A" Or Items(1) <> "B" Or Items(2) <> "C" Or Items(3) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
    Keys = dc1.Keys: If Keys(0) <> "a" Or Keys(1) <> Empty Or Keys(2) <> "c" Or Keys(3) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)

currentTest = "[FA2C-1] Function Add2 must be chainable"
    Set dc1 = New DictCollection: dc1.Add2("B").Add2("A", "a", 1).Add2("C", "c", , 1).Add2 Empty
    Items = dc1.Items: If Items(0) <> "A" Or Items(1) <> "B" Or Items(2) <> "C" Or Items(3) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
    Keys = dc1.Keys: If Keys(0) <> "a" Or Keys(1) <> Empty Or Keys(2) <> "c" Or Keys(3) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)

currentTest = "[FIC-1] Function Insert must be chainable"
    Set dc1 = New DictCollection: dc1.Insert("B", 0).Insert("A", 0, "a").Insert("C", 2, "c").Insert Empty
    Items = dc1.Items: If Items(0) <> "A" Or Items(1) <> "B" Or Items(2) <> "C" Or Items(3) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
    Keys = dc1.Keys: If Keys(0) <> "a" Or Keys(1) <> Empty Or Keys(2) <> "c" Or Keys(3) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)

currentTest = "[FRC-1] Function Remove must be chainable"
    Set dc1 = New DictCollection: dc1.Add "a", 1: dc1.Add , 2
    dc1.Remove("a").Remove(0).Remove -1
    If dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)

currentTest = "[FRAC-1] Function RemoveAll must be chainable"
    Set dc1 = New DictCollection: dc1.Add "a", 1: dc1.Add , 2
    dc1.RemoveAll().Add , 1
    If dc1.Count <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)

currentTest = "[FMIC-1] Function Move must be chainable"
    Set dc1 = New DictCollection: dc1.Add "c", "C": dc1.Add , "B": dc1.Add "a", "A"
    dc1.Move(2, 0).Move(1, 3).Move 3, 2
    Items = dc1.Items: If Items(0) <> "A" Or Items(1) <> "B" Or Items(2) <> "C" Or Items(3) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
    Keys = dc1.Keys: If Keys(0) <> "a" Or Keys(1) <> Empty Or Keys(2) <> "c" Or Keys(3) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)


currentTest = "[DIA-1] Direct index access must work with standard items and objects"
    Set dc1 = New DictCollection: dc1.NonExistingValue = Null: Set dc2 = New DictCollection: dc2.DefaultValueEnabled = False
    dc1.Add , "A": dc1.Add , dc2
    If dc1.ItemAt(-1) <> Null Or dc1.ItemAt(2) <> Null Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.ItemAt(0) <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)
    If VarType(dc1.ItemAt(1)) <> VbVarType.vbObject Or Not (dc1.ItemAt(1) Is dc2) Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[DIA-2] Direct index access must throw correct errors with .ThrowErros=True"
    dc1.ThrowErrors = True: On Error Resume Next
    Err.Clear: testdata = dc1.ItemAt(-1): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: testdata = dc1.ItemAt(2): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: Set testdata = dc1.ItemAt(1): If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail

currentTest = "[CIP-1] Changing items must work with standard items and objects"
    Set dc1 = New DictCollection: dc1.NonExistingValue = Null: Set dc2 = New DictCollection: dc2.DefaultValueEnabled = False
    dc1(-1) = "X" ' must do nothing
    If dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1(2) = "A" ' dc1.Items = Array (Empty,Empty,"A")
    If dc1.Count <> 3 Or dc1(0) <> Empty Or dc1(1) <> Empty Or dc1(2) <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)
    Set dc1(1) = dc2 ' dc1.Items = Array (Empty,dc2,"A")
    If dc1.Count <> 3 Or dc1(0) <> Empty Or Not (dc1(1) Is dc2) Or dc1(2) <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1(0) = -1 ' dc1.Items = Array (-1,dc2,"A")
    If dc1.Count <> 3 Or dc1(0) <> -1 Or Not (dc1(1) Is dc2) Or dc1(2) <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1(1) = 1 ' dc1.Items = Array (-1,1,"A")
    If dc1.Count <> 3 Or dc1(0) <> -1 Or dc1(1) <> 1 Or dc1(2) <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)

currentTest = "[IOK-1] IndexOfKey must return -1 when key is not found or key"
    Set dc1 = New DictCollection
    On Error Resume Next
    testdata = Array(CInt(1), CLng(2), CSng(3), CDbl(4), CCur(5), CDec(6), CDate(7), "", "a")
    For i = 0 To UBound(testdata)
        Err.Clear: val = dc1.IndexOfKey(testdata(i)): If Err.Number <> 0 And val <> -1 Then logError DebugPrint, errors, ("-> " & currentTest & ": nonexisting keys of type " & TypeName(testdata(i)) & " must return -1 and not throw errors")
    Next
    On Error GoTo Fail
currentTest = "[IOK-2] IndexOfKey must throw correct errors with .ThrowErrors=True"
    On Error Resume Next
    dc1.ThrowErrors = True: Set dc2 = New DictCollection: dc2.DefaultValueEnabled = False
    Err.Clear: val = dc1.IndexOfKey(Null): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest & ": empty DictCollection must throw Error 5 'Invalid procedure call or argument'")
    dc1.Add , "x"
    testdata = Array(Null, CVErr(11), False, dc2, Nothing)
    For i = 0 To UBound(testdata)
        Err.Clear: val = dc1.IndexOfKey(testdata(i)): If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest & ": keys of type " & TypeName(testdata(i)) & " must throw Error 13 'Type mismatch'")
    Next
    testdata = Array(CInt(1), CLng(2), CSng(3), CDbl(4), CCur(5), CDec(6), CDate(7), "a", "NONEXISTING")
    For i = 0 To UBound(testdata)
        Err.Clear: val = dc1.IndexOfKey(testdata(i)): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest & ": nonexisting keys must throw Error 5 'Invalid procedure call or argument'")
    Next
    On Error GoTo Fail
currentTest = "[IOK-3] IndexOfKey must return correct item indexes"
    dc1.RemoveAll
    testdata = Array(CInt(1), CLng(2), CSng(3), CDbl(4), CCur(5), CDec(6), #1/1/2000 11:00:00 AM#, "a")
    For i = 0 To UBound(testdata): dc1.Add CStr(testdata(i)), i: Next ' item = i
    On Error Resume Next
    For i = 0 To UBound(testdata)
        Err.Clear: val = dc1.IndexOfKey(testdata(i)): If val <> i Or Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": existing key of type " & TypeName(testdata(i)) & " must return correct index and not throw errors")
    Next
    On Error GoTo Fail: Err.Clear
currentTest = "[IOK-4] IndexOfKey must not throw errors with .EmulateDictionary=True"
    Set dc1 = New DictCollection: dc1.EmulateDictionary = True: Set dc2 = New DictCollection: dc2.DefaultValueEnabled = False
    If dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error Resume Next
    testdata = Array(CInt(1), CLng(2), CSng(3), CDbl(4), CCur(5), CDec(6), CDate(7), Null, CVErr(11), False, dc2, Nothing, "a", "")
    For i = 0 To UBound(testdata)
        dc1.ThrowErrors = True
        Err.Clear: val = dc1.IndexOfKey(testdata(i)): If val <> -1 Or Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": nonexisting keys must return -1 and not throw errors with .ThrowErrors=True")
        dc1.ThrowErrors = False
        Err.Clear: val = dc1.IndexOfKey(testdata(i)): If val <> -1 Or Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": nonexisting keys must return -1 and not throw errors with .ThrowErrors=False")
    Next
    On Error GoTo Fail: Err.Clear
currentTest = "[IOK-5] IndexOfKey must throw correct errors with .EmulateCollection=True"
    Set dc1 = New DictCollection: dc1.EmulateCollection = True: Set dc2 = New DictCollection: dc2.DefaultValueEnabled = False
    If dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error Resume Next
    ' empty DictCollection throws Error 5 'invalid procedure or call' to stay compatible with Collection
    testdata = Array(CInt(1), CLng(2), CSng(3), CDbl(4), CCur(5), CDec(6), CDate(7), Null, CVErr(11), False, dc2, Nothing, Empty, UtilGetMissingValue)
    For i = 0 To UBound(testdata)
           Err.Clear: val = dc1.IndexOfKey(testdata(i)): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest & ": keys of type " & TypeName(testdata(i)) & " must throw error 5 in empty DictCollection")
    Next
    dc1.Add , 1
    ' filled DictCollection throws Error 13 Type mismatch to stay of compatible with Collection
    testdata = Array(CInt(1), CLng(2), CSng(1.5), CDbl(2.5), CCur(5), CDec(6), CDate(7), Null, CVErr(11), False, dc2, Nothing, Empty, UtilGetMissingValue)
    For i = 0 To UBound(testdata)
           Err.Clear: val = dc1.IndexOfKey(testdata(i)): If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest & ": nonexisting keys throw error 13 in filled DictCollection")
    Next
    On Error GoTo Fail: Err.Clear

currentTest = "[KOIA-1] KeyOfItemAt must return Empty if item has no key or index does not exist"
    Set dc1 = New DictCollection
    On Error Resume Next
    Err.Clear: val = dc1.KeyOfItemAt(-1): If dc1.Count <> 0 Or Err.Number <> 0 Or val <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.KeyOfItemAt(0): If dc1.Count <> 0 Or Err.Number <> 0 Or val <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.KeyOfItemAt(25): If dc1.Count <> 0 Or Err.Number <> 0 Or val <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear
currentTest = "[KOIA-2] KeyOfItemAt must throw correct errors with .ThrowErrors=True"
    Set dc1 = New DictCollection: dc1.ThrowErrors = True
    On Error Resume Next
    ' empty DictCollection throws Error 5 while filled DictCollection throws Error 9 to stay of compatible with Collection
    Err.Clear: val = dc1.KeyOfItemAt(-1): If dc1.Count <> 0 Or Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.KeyOfItemAt(0): If dc1.Count <> 0 Or Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.KeyOfItemAt(25): If dc1.Count <> 0 Or Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Add , "itemA": dc1.Add "b", "itemB": Set dc2 = New DictCollection
    Err.Clear: dc1.KeyOfItemAt(0) = "": Keys = dc1.Keys: If Err.Number <> 0 Or dc1.Count <> 2 Or Keys(0) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest & ": setting and items key to none must not throw errors")
    Err.Clear: dc1.KeyOfItemAt(0) = CDbl(1.01): Keys = dc1.Keys: If Err.Number <> 0 Or dc1.Count <> 2 Or Keys(0) <> CStr(CDbl(1.01)) Then logError DebugPrint, errors, ("-> " & currentTest & ": setting and items key to a string-convertable data type must not throw errors")
    Err.Clear: dc1.KeyOfItemAt(0) = CDate(1): Keys = dc1.Keys: If Err.Number <> 0 Or dc1.Count <> 2 Or Keys(0) <> CStr(CDate(1)) Then logError DebugPrint, errors, ("-> " & currentTest & ": setting and items key to a string-convertable data type must not throw errors")
    Err.Clear: dc1.KeyOfItemAt(0) = Empty: Keys = dc1.Keys: If Err.Number <> 0 Or dc1.Count <> 2 Or Keys(0) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest & ": setting and items key to none must not throw errors")
    Err.Clear: dc1.KeyOfItemAt(1) = dc2: Keys = dc1.Keys: If Err.Number <> 13 Or dc1.Count <> 2 Or Keys(1) <> "b" Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.CompareMode = vbTextCompare
    Err.Clear: dc1.KeyOfItemAt(1) = "B": Keys = dc1.Keys: If Err.Number <> 0 Or dc1.Count <> 2 Or Keys(1) <> "B" Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.KeyOfItemAt(-1): If Err.Number <> 9 Or dc1.Count <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.KeyOfItemAt(2): If Err.Number <> 9 Or dc1.Count <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.ZeroBasedIndex = False
    Err.Clear: val = dc1.KeyOfItemAt(0): If Err.Number <> 9 Or dc1.Count <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.KeyOfItemAt(3): If Err.Number <> 9 Or dc1.Count <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear
currentTest = "[KOIA-3] KeyOfItemAt must get and set correct keys and set existing keys to none if necessary"
    Set dc1 = New DictCollection: dc1.Add , "Item1": dc1.Add 1, "Item2": dc1.Add "a", "Item3": dc1.Add "A", "Item4"
        If dc1.Count <> 4 Or dc1.KeyOfItemAt(0) <> Empty Or dc1.KeyOfItemAt(1) <> CStr(1) Or dc1.KeyOfItemAt(2) <> "a" Or dc1.KeyOfItemAt(3) <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.KeyOfItemAt(0) = "1" ' this will set the key of the item at index=1 "Item2" to none because it was previously using the key "1"
        If dc1.Count <> 4 Or dc1.KeyOfItemAt(0) <> "1" Or dc1.KeyOfItemAt(1) <> Empty Or dc1.KeyOfItemAt(2) <> "a" Or dc1.KeyOfItemAt(3) <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc1.Item(0) <> "Item1" Or dc1.Item(1) <> "Item2" Or dc1.Item(2) <> "Item3" Or dc1.Item(3) <> "Item4" Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.KeyOfItemAt(1) = "2": dc1.KeyOfItemAt(2) = "3": dc1.KeyOfItemAt(3) = "4"
        If dc1.Count <> 4 Or dc1.KeyOfItemAt(0) <> "1" Or dc1.KeyOfItemAt(1) <> "2" Or dc1.KeyOfItemAt(2) <> "3" Or dc1.KeyOfItemAt(3) <> "4" Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc1.Item(0) <> "Item1" Or dc1.Item(1) <> "Item2" Or dc1.Item(2) <> "Item3" Or dc1.Item(3) <> "Item4" Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.KeyOfItemAt(5) = "6" ' inserts item at index=5 with key "6" and implicitly adds an Empty item at index=4
        If dc1.Count <> 6 Or dc1.KeyOfItemAt(0) <> "1" Or dc1.KeyOfItemAt(1) <> "2" Or dc1.KeyOfItemAt(2) <> "3" Or dc1.KeyOfItemAt(3) <> "4" Or dc1.KeyOfItemAt(4) <> Empty Or dc1.KeyOfItemAt(5) <> "6" Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc1.Item(0) <> "Item1" Or dc1.Item(1) <> "Item2" Or dc1.Item(2) <> "Item3" Or dc1.Item(3) <> "Item4" Or dc1.Item(4) <> Empty Or dc1.Item(5) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.KeyOfItemAt(5) = "1" ' this will set the key of the item at index=0 "Item1" to none because it was previously using the key "1"
        If dc1.Count <> 6 Or dc1.KeyOfItemAt(0) <> Empty Or dc1.KeyOfItemAt(1) <> "2" Or dc1.KeyOfItemAt(2) <> "3" Or dc1.KeyOfItemAt(3) <> "4" Or dc1.KeyOfItemAt(4) <> Empty Or dc1.KeyOfItemAt(5) <> "1" Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc1.Item(0) <> "Item1" Or dc1.Item(1) <> "Item2" Or dc1.Item(2) <> "Item3" Or dc1.Item(3) <> "Item4" Or dc1.Item(4) <> Empty Or dc1.Item(5) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.KeyOfItemAt(0) = "1" ' this will set the key of the item at index=5 to none because it was previously using the key "1"
        If dc1.Count <> 6 Or dc1.KeyOfItemAt(0) <> "1" Or dc1.KeyOfItemAt(1) <> "2" Or dc1.KeyOfItemAt(2) <> "3" Or dc1.KeyOfItemAt(3) <> "4" Or dc1.KeyOfItemAt(4) <> Empty Or dc1.KeyOfItemAt(5) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc1.Item(0) <> "Item1" Or dc1.Item(1) <> "Item2" Or dc1.Item(2) <> "Item3" Or dc1.Item(3) <> "Item4" Or dc1.Item(4) <> Empty Or dc1.Item(5) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
        Keys = dc1.Keys: Items = dc1.Items
        If Keys(0) <> "1" Or Keys(1) <> "2" Or Keys(2) <> "3" Or Keys(3) <> "4" Or Keys(4) <> Empty Or Keys(5) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
        If Items(0) <> "Item1" Or Items(1) <> "Item2" Or Items(2) <> "Item3" Or Items(3) <> "Item4" Or Items(4) <> Empty Or Items(5) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear
currentTest = "[KOIA-4] KeyOfItemAt must work with .EmulateDictionary=True"
    Set dc1 = New DictCollection: dc1.EmulateDictionary = True: Set dc2 = New DictCollection: dc2.DefaultValueEnabled = False
    testdata = Array(UtilGetMissingValue, CVErr(5), "", Null, "a", "B", CInt(1), CDbl(2), CDate(3), dc2)
    For i = 0 To UBound(testdata): dc1.Add testdata(i), "Item" & i: Next
    For i = 0 To UBound(testdata)
        If IsObject(testdata(i)) Then
            Set val = dc1.KeyOfItemAt(i)
            If TypeName(val) <> TypeName(testdata(i)) Then
                logError DebugPrint, errors, ("-> " & currentTest & ": key at " & i & " must be of type " & TypeName(testdata(i)))
            Else
                If Not val Is testdata(i) Then logError DebugPrint, errors, ("-> " & currentTest & ": key at " & i & " must be same object as testdata")
            End If
        Else
            val = dc1.KeyOfItemAt(i)
            If TypeName(val) <> TypeName(testdata(i)) Then
                logError DebugPrint, errors, ("-> " & currentTest & ": key at " & i & " must be of type " & TypeName(testdata(i)))
            Else
                If val <> testdata(i) Then logError DebugPrint, errors, ("-> " & currentTest & ": key at " & i & " must be" & CStr(testdata(i)))
            End If
        End If
    Next
    On Error Resume Next
    ' must throw Error 457 'Key is already associated with an item'
    Err.Clear: dc1.KeyOfItemAt(0) = Empty: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & ": must throw Error 457 'Key is already associated with an item'")
    Err.Clear: dc1.KeyOfItemAt(0) = CSng(1): If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & ": must throw Error 457 'Key is already associated with an item'")
    Err.Clear: dc1.KeyOfItemAt(0) = dc2: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & ": must throw Error 457 'Key is already associated with an item'")
    Err.Clear: dc1.KeyOfItemAt(0) = Null: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & ": must throw Error 457 'Key is already associated with an item'")
    Err.Clear: dc1.KeyOfItemAt(0) = testdata(0): If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": must throw Error 457 'Key is already associated with an item'")
    ' must throw Error 9 'Subscript out of range'
    Err.Clear: val = dc1.KeyOfItemAt(-1): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest & ": must throw Error 9 'Subscript out of range'")
    Err.Clear: val = dc1.KeyOfItemAt(UBound(testdata) + 1): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest & currentTest & ": must throw Error 9 'Subscript out of range'")
    Err.Clear: dc1.KeyOfItemAt(-1) = "x": If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest & currentTest & ": must throw Error 9 'Subscript out of range'")
    Err.Clear: dc1.KeyOfItemAt(UBound(testdata) + 1) = "x": If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest & currentTest & ": must throw Error 9 'Subscript out of range'")
    dc1.ThrowErrors = False
    ' remove first key
    Err.Clear: dc1.KeyOfItemAt(0) = testdata(UBound(testdata)) ' set key of item at index=0 to key of last item
        If Err.Number <> 0 Or dc1.Count <> UBound(testdata) + 1 Or dc1.KeyCount <> UBound(testdata) Then logError DebugPrint, errors, ("-> " & currentTest & ": after removing first key must have .KeyCount=" & UBound(testdata))
    ' remove second key
    Err.Clear: dc1.KeyOfItemAt(0) = "a" ' set key of item at index=0 to key "a" (sets the key of the item that previously used that key to none)
        If Err.Number <> 0 Or dc1.Count <> UBound(testdata) + 1 Or dc1.KeyCount <> UBound(testdata) - 1 Or dc1.KeyOfItemAt(0) <> "a" Then logError DebugPrint, errors, ("-> " & currentTest & ": after removing second key must have .KeyCount=" & UBound(testdata) - 1)
    ' remove third key
    dc1.CompareMode = vbTextCompare
    Err.Clear: dc1.KeyOfItemAt(0) = "b" ' set key of item at index=0 to key "b" (sets the key of the item that previously used that key to none)
        If Err.Number <> 0 Or dc1.Count <> UBound(testdata) + 1 Or dc1.KeyCount <> UBound(testdata) - 2 Or dc1.KeyOfItemAt(0) <> "b" Then logError DebugPrint, errors, ("-> " & currentTest & ": after removing third key must have .KeyCount=" & UBound(testdata) - 2)
currentTest = "[KOIA-5] KeyOfItemAt must work with .EmulateCollection=True"
    Set dc1 = New DictCollection: dc1.EmulateCollection = True: Set dc2 = New DictCollection: dc2.DefaultValueEnabled = False
    On Error Resume Next
    ' must throw Error 5 'Invalid procedure call or argument' with empty Collection
    Err.Clear: val = dc1.KeyOfItemAt(-1): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.KeyOfItemAt(0): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.KeyOfItemAt(1): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Add2 "Item1", "" ' key = "" is supported by Collection
    ' must throw Error 9 'Subscript out of range'
    Err.Clear: val = dc1.KeyOfItemAt(-1): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.KeyOfItemAt(0): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.KeyOfItemAt(9): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.KeyOfItemAt(1): If Err.Number <> 0 Or val <> "" Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Add2 "Item2"
    If dc1.Count <> 2 Or dc1.KeyCount <> 1 Or dc1.KeyOfItemAt(1) <> "" Or dc1.KeyOfItemAt(2) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Add2 "Item3", "a"
    ' dc1.Items = Array("Item1","Item2","Item3") dc1.Keys = Array("",Empty,"a")
        If dc1.Count <> 3 Or dc1.KeyCount <> 2 Or dc1.KeyOfItemAt(1) <> "" Or dc1.KeyOfItemAt(2) <> Empty Or dc1.KeyOfItemAt(3) <> "a" Then logError DebugPrint, errors, ("-> " & currentTest)
    ' must throw Error 457 'Key is already associated with an item
    Err.Clear: dc1.Add2 "Item4", "A": If Err.Number <> 457 Or dc1.Count <> 3 Or dc1.KeyCount <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: dc1.Add2 "Item4", "": If Err.Number <> 457 Or dc1.Count <> 3 Or dc1.KeyCount <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    ' must throw Error 13 'Type mismatch'
    Err.Clear: dc1.KeyOfItemAt(1) = CInt(1): If Err.Number <> 13 Or dc1.Count <> 3 Or dc1.KeyCount <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: dc1.KeyOfItemAt(1) = CDate(1): If Err.Number <> 13 Or dc1.Count <> 3 Or dc1.KeyCount <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: dc1.KeyOfItemAt(1) = dc2: If Err.Number <> 13 Or dc1.Count <> 3 Or dc1.KeyCount <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: dc1.KeyOfItemAt(1) = UtilGetMissingValue: If Err.Number <> 13 Or dc1.Count <> 3 Or dc1.KeyCount <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.ThrowErrors = False
    ' test assigning of keys
    Err.Clear: dc1.KeyOfItemAt(1) = "a" ' will set key of "Item3" to Empty because it already uses the key "a"
        If Err.Number <> 0 Or dc1.KeyCount <> 1 Or dc1.KeyOfItemAt(1) <> "a" Or dc1.KeyOfItemAt(2) <> Empty Or dc1.KeyOfItemAt(3) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: dc1.KeyOfItemAt(5) = "A" ' will set key of "Item1" to Empty because it already uses the key "a" and insert two Empty items at end (index=4,5)
    ' dc1.Items = Array("Item1","Item2","Item3",Empty,Empty) dc1.Keys = Array(Empty,Empty,Empty,Empty,"A")
        If Err.Number <> 0 Or dc1.KeyCount <> 1 Or dc1.KeyOfItemAt(1) <> Empty Or dc1.KeyOfItemAt(2) <> Empty Or dc1.KeyOfItemAt(3) <> Empty Or dc1.KeyOfItemAt(4) <> Empty Or dc1.KeyOfItemAt(5) <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)
        Items = dc1.Items
        If Items(0) <> "Item1" Or Items(1) <> "Item2" Or Items(2) <> "Item3" Or Items(3) <> Empty Or Items(4) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
    ' must also work with ZeroBasedIndex=True
    dc1.ZeroBasedIndex = True
        If Err.Number <> 0 Or dc1.KeyCount <> 1 Or dc1.KeyOfItemAt(0) <> Empty Or dc1.KeyOfItemAt(1) <> Empty Or dc1.KeyOfItemAt(2) <> Empty Or dc1.KeyOfItemAt(3) <> Empty Or dc1.KeyOfItemAt(4) <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear

currentTest = "[SK-1] Sorted Keys must return empty array if there are no keys"
    Set dc1 = New DictCollection
    Keys = dc1.SortedKeys
    If UBound(Keys) <> -1 Then logError DebugPrint, errors, ("-> " & currentTest)


    For i = 1 To 2
        If i = 1 Then
currentTest = "[SK-2] Sorted Keys must return keys in correct sorted order with binary compare and eager sorting"
            Set dc1 = New DictCollection: dc1.CompareMode = vbBinaryCompare: dc1.LazySorting = False
        Else
currentTest = "[SK-3] Sorted Keys must return keys in correct sorted order with binary compare and lazy sorting"
            Set dc1 = New DictCollection: dc1.CompareMode = vbBinaryCompare: dc1.LazySorting = True
        End If
        dc1(ChrW(9702)) = "BULLET5"                     ' key1 ?
        dc1(ChrW(224)) = "SMALL_LETTER_A_WITH_ACCENT"   ' key2 ‡
        dc1(ChrW(183)) = "BULLET1"                      ' key3 ∑
        dc1.Add , "ItemWithoutKey"
        dc1(ChrW(97)) = "SMALL_LETTER_A"                ' key4 a
        dc1(ChrW(65)) = "BIG_LETTER_A"                  ' key5 A
        dc1(ChrW(229)) = "SMALL_LETTER_A_WITH_CIRCLE"   ' key6 Â
        dc1(ChrW(9679)) = "BULLET4"                     ' key7 ?
        dc1(ChrW(196)) = "BIG_LETTER_A_WITH_UMLAUT"     ' key8 ƒ
        dc1(ChrW(8729)) = "BULLET3"                     ' key9 ∑
        dc1(ChrW(197)) = "BIG_LETTER_A_WITH_CIRCLE"     ' key10 ≈
        dc1(ChrW(8226)) = "BULLET2"                     ' key11 ï
        dc1(ChrW(192)) = "BIG_LETTER_A_WITH_ACCENT"     ' key12 ¿
        dc1(ChrW(228)) = "SMALL_LETTER_A_WITH_UMLAUT"   ' key13 ‰
        Keys = dc1.SortedKeys
        If UBound(Keys) <> 12 Then logError DebugPrint, errors, ("-> " & currentTest)
        ' test against static list
        If Keys(0) <> ChrW(65) Or Keys(1) <> ChrW(97) Or Keys(2) <> ChrW(183) Or Keys(3) <> ChrW(192) Or Keys(4) <> ChrW(196) Then logError DebugPrint, errors, ("-> " & currentTest)
        If Keys(5) <> ChrW(197) Or Keys(6) <> ChrW(224) Or Keys(7) <> ChrW(228) Or Keys(8) <> ChrW(229) Or Keys(9) <> ChrW(8226) Then logError DebugPrint, errors, ("-> " & currentTest)
        If Keys(10) <> ChrW(8729) Or Keys(11) <> ChrW(9679) Or Keys(12) <> ChrW(9702) Then logError DebugPrint, errors, ("-> " & currentTest)
        ' test against reference sort function that emulates binary string sorting
        testdata = Array() ' create empty array
        For j = 0 To UBound(Keys): dc1.UtilAddArrayValue testdata, AscW(Keys(j)): Next ' add all keys as numbers (unicode code points)
        dc1.UtilSortArray testdata, 0, UBound(testdata) ' sort the number array (emulating binary string sort)
        For j = 0 To UBound(Keys)
            If AscW(Keys(j)) <> testdata(j) Then logError DebugPrint, errors, ("-> " & currentTest)
        Next
    Next
    
    For i = 1 To 2
        If i = 1 Then
currentTest = "[SK-4] Sorted Keys must return keys in correct sorted order with case insensitive text compare and eager sorting"
            Set dc1 = New DictCollection: dc1.CompareMode = vbTextCompare: dc1.LazySorting = False
        Else
currentTest = "[SK-5] Sorted Keys must return keys in correct sorted order with case insensitive text compare and lazy sorting"
            Set dc1 = New DictCollection: dc1.CompareMode = vbTextCompare: dc1.LazySorting = True
        End If
        dc1(ChrW(9702)) = "BULLET5"                     ' key1 ?
        dc1(ChrW(224)) = "SMALL_LETTER_A_WITH_ACCENT"   ' key2 ‡
        dc1(ChrW(183)) = "BULLET1"                      ' key3 ∑
        dc1(ChrW(97)) = "SMALL_LETTER_A"                ' key4 a
        dc1(ChrW(65)) = "BIG_LETTER_A"                  ' key4 A -> overwrites key4
        dc1(ChrW(229)) = "SMALL_LETTER_A_WITH_CIRCLE"   ' key5 Â
        dc1(ChrW(9679)) = "BULLET4"                     ' key6 ?
        dc1(ChrW(196)) = "BIG_LETTER_A_WITH_UMLAUT"     ' key7 ƒ
        dc1(ChrW(8729)) = "BULLET3"                     ' key8 ∑
        dc1(ChrW(197)) = "BIG_LETTER_A_WITH_CIRCLE"     ' key5 ≈ -> overwrites key5
        dc1(ChrW(8226)) = "BULLET2"                     ' key9 ï
        dc1(ChrW(192)) = "BIG_LETTER_A_WITH_ACCENT"     ' key2 ¿ -> overwrites key2
        dc1(ChrW(228)) = "SMALL_LETTER_A_WITH_UMLAUT"   ' key7 ‰ -> overwrites key7
        Keys = dc1.SortedKeys
        If UBound(Keys) <> 8 Then logError DebugPrint, errors, ("-> " & currentTest)
        ' test against static list
        If Keys(0) <> ChrW(8729) Or Keys(1) <> ChrW(9679) Or Keys(2) <> ChrW(9702) Or Keys(3) <> ChrW(183) Or Keys(4) <> ChrW(8226) Then logError DebugPrint, errors, ("-> " & currentTest)
        If Keys(5) <> ChrW(65) Or Keys(6) <> ChrW(192) Or Keys(7) <> ChrW(228) Or Keys(8) <> ChrW(197) Then logError DebugPrint, errors, ("-> " & currentTest)
        ' test against reference sort function that emulates case insensitive string sorting
        testdata = Array() ' create empty array
        For j = 0 To UBound(Keys): dc1.UtilAddArrayValue testdata, Keys(j): Next ' add all keys
        dc1.UtilSortStringArray testdata, 0, UBound(testdata), vbTextCompare ' sort the array (emulating case insensitive string sort)
        For j = 0 To UBound(Keys)
            If AscW(Keys(j)) <> AscW(testdata(j)) Then logError DebugPrint, errors, ("-> " & currentTest)
        Next
    Next

currentTest = "[PADC-1] AsDC must return nonexisting DictCollection if index or key does not exist or item is not DictCollection"
    Set dc1 = New DictCollection
    ' weird: dc1.AsDC(-1).CollectionType <> 0 will evaluate to true if .CollectionType = 0
    val = dc1.AsDC(-1).CollectionType: If val <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    val = dc1.AsDC("").CollectionType: If val <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.AsDC("a") <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Add "x", 1
    val = dc1.AsDC(-1).CollectionType: If val <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    val = dc1.AsDC(1).CollectionType: If val <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    val = dc1.AsDC("a").CollectionType: If val <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    val = dc1.AsDC("x").CollectionType: If val <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[PADC-2] AsDC must return throw correct errors if index or key does not exist or existing item is not DictCollection"
    Set dc1 = New DictCollection: dc1.ThrowErrors = True
    ' errors are inherited from internalGetItem()
    On Error Resume Next
    ' empty DictCollection
    Err.Clear: val = dc1.AsDC(-1): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.AsDC(0): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.AsDC("a"): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Add "x", 1 ' filled DictCollection
    Err.Clear: val = dc1.AsDC(-1): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.AsDC("a"): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.AsDC(0): If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.AsDC("x"): If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear
currentTest = "[PADC-3] AsDC must return correct DictCollection"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection: dc2.Item(0) = "X"
    dc1.Add "x", dc2
    If dc1.AsDC("x").Item(0) <> "X" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.AsDC(0).Item(0) <> "X" Then logError DebugPrint, errors, ("-> " & currentTest)
    
currentTest = "[PIA-1] Get ItemAt must retrieve correct items"
    Set dc1 = New DictCollection: dc1.Add , 1: dc1.Add , 2: dc1.Add , 3
    If dc1.ItemAt(0) <> 1 Or dc1.ItemAt(1) <> 2 Or dc1.ItemAt(2) <> 3 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[PIA-2] Get ItemAt must return .NonExistingValue for nonexisting indexes"
    If dc1.ItemAt(-1) <> dc1.NonExistingValue Or dc1.ItemAt(3) <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.NonExistingValue = 0
    If dc1.ItemAt(-1) <> dc1.NonExistingValue Or dc1.ItemAt(3) <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[PIA-3] Get ItemAt must throw correct errors or return .NonExistingValue"
    Set dc1 = New DictCollection: dc1.ThrowErrors = True
    On Error Resume Next
    ' Empty DictCollection
    Err.Clear: val = dc1.ItemAt(-1): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.ItemAt(0): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Add "", 1 ' filled DictCollection
    Err.Clear: val = dc1.ItemAt(-1): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.ItemAt(1): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.ThrowErrors = False ' return nonexisting value
    Err.Clear: val = dc1.ItemAt(-1): If Err.Number <> 0 Or val <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.ItemAt(1): If Err.Number <> 0 Or val <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.NonExistingValue = "x" ' return modified nonexisting value
    Err.Clear: val = dc1.ItemAt(2): If Err.Number <> 0 Or val <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.ItemAt(3): If Err.Number <> 0 Or val <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear
currentTest = "[PIA-4] Let/Set ItemAt must assign or add correct items"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.ItemAt(1) = 1 ' dc1.Items = Array(Empty,1)
        Items = dc1.Items: If dc1.Count <> 2 Or Items(0) <> Empty Or Items(1) <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.ItemAt(-1) = 2 ' must do nothing
        Items = dc1.Items: If dc1.Count <> 2 Or Items(0) <> Empty Or Items(1) <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.ItemAt(0) = "0" ' dc1.Items = Array("0",1)
    dc1.ItemAt(2) = dc2: Set dc1.ItemAt(3) = dc2 ' dc1.Items = Array("0",1,dc2,dc2)
        Items = dc1.Items: If dc1.Count <> 4 Or Items(0) <> "0" Or Items(1) <> 1 Or Not IsObject(Items(2)) Or Not IsObject(Items(3)) Then logError DebugPrint, errors, ("-> " & currentTest)
        If Not IsObject(dc1.ItemAt(2)) Or TypeName(dc1.ItemAt(2)) <> "DictCollection" Then
            UtilAddArrayValue errors, ("-> " & currentTest)
        Else
            If Not dc1.ItemAt(2) Is dc2 Then logError DebugPrint, errors, ("-> " & currentTest)
        End If
        If Not IsObject(dc1.ItemAt(3)) Or TypeName(dc1.ItemAt(3)) <> "DictCollection" Then
            UtilAddArrayValue errors, ("-> " & currentTest)
        Else
            If Not dc1.ItemAt(3) Is dc2 Then logError DebugPrint, errors, ("-> " & currentTest)
        End If
currentTest = "[PIA-5] Let/Set ItemAt must throw correct errors"
    Set dc1 = New DictCollection: dc1.ThrowErrors = True: Set dc2 = New DictCollection
    On Error Resume Next
    Err.Clear: dc1.ItemAt(-1) = "a": If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: Set dc1.ItemAt(-1) = dc2: If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: dc1.ItemAt(0) = "a": If Err.Number <> 9 Or dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: dc1.ItemAt(1) = dc2: If Err.Number <> 9 Or dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: Set dc1.ItemAt(0) = dc2: If Err.Number <> 9 Or dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: Set dc1.ItemAt(1) = dc2: If Err.Number <> 9 Or dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear
currentTest = "[PIA-6] Get/Let/Set ItemAt must work with .ZeroBasedIndex=False"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.ZeroBasedIndex = False
    On Error Resume Next
    dc1.ThrowErrors = True ' must fail
    Err.Clear: dc1.ItemAt(0) = "a": If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: Set dc1.ItemAt(0) = dc2: If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.ThrowErrors = False ' must work
    Err.Clear: dc1.ItemAt(1) = "a": If Err.Number <> 0 Or dc1.Count <> 1 Or dc1.ItemAt(1) <> "a" Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: Set dc1.ItemAt(2) = dc2: If Err.Number <> 0 Or dc1.Count <> 2 Or Not IsObject(dc1.ItemAt(2)) Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear
    
currentTest = "[PIO-1] Get ItemOf must retrieve correct items"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.Add "a", 1: dc1.Add "b", 2:  dc1.Add "c", dc2
    If dc1.ItemOf("a") <> 1 Or dc1.ItemOf("b") <> 2 Or Not IsObject(dc1.ItemOf("c")) Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[PIO-2] Get ItemOf must throw correct errors or return .NonExistingValue"
    Set dc1 = New DictCollection: dc1.ThrowErrors = True
    On Error Resume Next
    ' empty DictCollection
    Err.Clear: val = dc1.ItemOf("a"): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.ItemOf(""): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Add "b", 1 ' filled DictCollection
    Err.Clear: val = dc1.ItemOf("a"): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.ItemOf(""): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.ThrowErrors = False ' return nonexisting value
    Err.Clear: val = dc1.ItemOf("a"): If Err.Number <> 0 Or val <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.ItemOf(""): If Err.Number <> 0 Or val <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.NonExistingValue = -1 ' return modified nonexisting value
    Err.Clear: val = dc1.ItemOf("x"): If Err.Number <> 0 Or val <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.ItemOf(""): If Err.Number <> 0 Or val <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear
currentTest = "[PIO-3] Let/Set ItemOf must assign correct items"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.ItemOf("a") = 1 ' dc1.Items = Array(1) dc1.Keys = Array("a")
        Items = dc1.Items: Keys = dc1.Keys: If dc1.Count <> 1 Or dc1.KeyCount <> 1 Or Items(0) <> 1 Or Keys(0) <> "a" Or dc1.ItemOf("a") <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.ItemOf("") = 2 ' must do nothing
        Items = dc1.Items: Keys = dc1.Keys: If dc1.Count <> 1 Or dc1.KeyCount <> 1 Or Items(0) <> 1 Or Keys(0) <> "a" Or dc1.ItemOf("a") <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.ItemOf("b") = "2" ' dc1.Items = Array(1,"2") dc1.Keys = Array("a","b")
    dc1.ItemOf("c") = dc2: Set dc1.ItemOf("d") = dc2 ' dc1.Items = Array(1,"2",dc2,dc2) dc1.Keys = Array("a","b","c","d")
        Items = dc1.Items: Keys = dc1.Keys
        If dc1.Count <> 4 Or Items(0) <> 1 Or Items(1) <> "2" Or Not IsObject(Items(2)) Or Not IsObject(Items(3)) Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc1.KeyCount <> 4 Or Keys(0) <> "a" Or Keys(1) <> "b" Or Keys(2) <> "c" Or Keys(3) <> "d" Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc1.ItemOf("a") <> 1 Or dc1.ItemOf("b") <> "2" Then logError DebugPrint, errors, ("-> " & currentTest)
        If Not IsObject(dc1.ItemOf("c")) Or TypeName(dc1.ItemOf("c")) <> "DictCollection" Then
            UtilAddArrayValue errors, ("-> " & currentTest)
        Else
            If Not dc1.ItemOf("c") Is dc2 Then logError DebugPrint, errors, ("-> " & currentTest)
        End If
        If Not IsObject(dc1.ItemOf("d")) Or TypeName(dc1.ItemOf("d")) <> "DictCollection" Then
            UtilAddArrayValue errors, ("-> " & currentTest)
        Else
            If Not dc1.ItemOf("d") Is dc2 Then logError DebugPrint, errors, ("-> " & currentTest)
        End If
currentTest = "[PIO-4] Let/Set ItemOf must throw correct errors"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    On Error Resume Next
    dc1.ThrowErrors = True ' empty DictCollection, must throw errors
    Err.Clear: dc1.ItemOf("a") = 1: If Err.Number <> 5 Or dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: Set dc1.ItemOf("b") = dc2: If Err.Number <> 5 Or dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Add "c", 3 ' filled array, must also throw errors
    Err.Clear: dc1.ItemOf("a") = 1: If Err.Number <> 5 Or dc1.Count <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: Set dc1.ItemOf("b") = dc2: If Err.Number <> 5 Or dc1.Count <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.ThrowErrors = False ' must work
    Err.Clear: dc1.ItemOf("a") = 1: If Err.Number <> 0 Or dc1.Count <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: Set dc1.ItemOf("b") = dc2: If Err.Number <> 0 Or dc1.Count <> 3 Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear
currentTest = "[PIO-5] Let/Set ItemOf must overwrite older keys with case insensitive text compare"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.CompareMode = vbTextCompare
    dc1.ItemOf("a") = 1
    If dc1.ItemOf("a") <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.ItemOf("A") = 2: Keys = dc1.Keys
    If dc1.KeyCount <> 1 Or dc1.Count <> 1 Or dc1.ItemOf("A") <> 2 And Keys(0) <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)
    Set dc1.ItemOf("a") = dc2: Keys = dc1.Keys
    If dc1.KeyCount <> 1 Or dc1.Count <> 1 Or Not IsObject(dc1.ItemOf("a")) And Keys(0) <> "a" Then logError DebugPrint, errors, ("-> " & currentTest)

    For i = 1 To 2
        If i = 1 Then
currentTest = "[EDT-1] Using exotic datatypes as items with eager sorting": Set dc1 = New DictCollection: dc1.LazySorting = False
        Else
currentTest = "[EDT-2] Using exotic datatypes as items with lazy sorting": Set dc1 = New DictCollection: dc1.LazySorting = True
        End If
        dc1.Add "", Nothing
        dc1.Add "", Null
        dc1.Add "", Empty
        dc1.Insert UtilGetMissingValue, 5 ' will fill in 2 Empty items at dc(3) and dc(4)
        ' dc1.Items = Array(Nothing, Null, Empty, Empty, Empty, Missing)
            If dc1.Count <> 6 Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1.Count must be 6")
            If Not dc1(0) Is Nothing Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1(0) must be Nothing")
            If dc1(1) <> Null Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1(1) must be Null")
            If dc1(2) <> Empty Or dc1(3) <> Empty Or dc1(4) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1(2)-dc1(4) must be Empty")
            If Not IsMissing(dc1(5)) Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1(5) must be Missing")
        Items = dc1.Items
            If Not Items(0) Is Nothing Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1.Items(0) must be Nothing")
            If Items(1) <> Null Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1.Items(1) must be Null")
            If Items(2) <> Empty Or Items(3) <> Empty Or Items(4) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1.Items(2)-dc1.Items(4) must be Empty")
            If Not IsMissing(Items(5)) Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1.Items(5) must be Missing")
        Items = dc1.Items(False)
            If Not Items(0) Is Nothing Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1.Items(false)(0) must be Nothing")
            If Items(1) <> Null Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1.Items(false)(1) must be Null")
            If Not IsMissing(Items(2)) Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1.Items(false)(2) must be Missing")
        dc1.Insert Nothing, 0, "first"
        dc1.Insert Nothing, , "last" ' will insert at end
        ' dc1.Items = Array(Nothing, Nothing, Null, Empty, Empty, Empty, Missing, Nothing)
        dc1.Remove 5: dc1.Remove 4: dc1.Remove 3 ' mind direction!
        ' dc1.Items = Array(Nothing, Nothing, Null, Missing, Nothing)
        dc1(1) = CVErr(11)
        ' now dc1.Items = Array(Nothing, CVErr(11), Null, Missing, Nothing)
            If dc1.Count <> 5 Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1.Count must be 5 after re-ordering")
            If Not dc1(0) Is Nothing Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1(0) must be Nothing after re-ordering")
            If Not IsError(dc1(1)) Then
                UtilAddArrayValue errors, ("-> " & currentTest & ": dc1(1) must be Error after re-ordering")
            Else
                If dc1(1) <> CVErr(11) Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1(1) must be CVErr(11) after re-ordering")
            End If
            If dc1(2) <> Null Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1(2) must be Null after re-ordering")
            If Not IsMissing(dc1(3)) Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1(3) must be Missing after re-ordering")
            If Not dc1(4) Is Nothing Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1(4) must be Nothing after re-ordering")
        dc1.Insert Array(1, 2), 1, "second"
        ' dc1.Items = Array(Nothing, Array, CVErr(11), Null, Missing, Nothing)
        Items = dc1.Items
            If UBound(Items) <> 5 Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1.Items must have 6 items after re-ordering and insert")
            If Not Items(0) Is Nothing Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1.Items(0) must be Nothing after re-ordering and insert")
            If Not IsArray(Items(1)) Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1.Items(1) must be Array after re-ordering and insert")
            If Items(2) <> CVErr(11) Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1.Items(2) must be Error 11 after re-ordering and insert")
            If Items(3) <> Null Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1.Items(3) must be Null after re-ordering and insert")
            If Not IsMissing(Items(4)) Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1.Items(4) must be Missing after re-ordering and insert")
            If Not Items(5) Is Nothing Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1.Items(5) must be Nothing after re-ordering and insert")
            If dc1.Count <> 6 Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1.Count must be 6 after re-ordering and insert")
            If Not dc1("first") Is Nothing Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1(""first"") must be Nothing after re-ordering and insert")
            If Not VBA.IsArray(dc1("second")) Then
                UtilAddArrayValue errors, ("-> " & currentTest & ": dc1(""second"") must be Array after re-ordering and insert")
            Else
                If Not dc1("second")(0) = 1 Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1(""second"")(0) must be 1 after re-ordering and insert")
                If Not dc1("second")(1) = 2 Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1(""second"")(0) must be 1 after re-ordering and insert")
            End If
            If Not dc1("last") Is Nothing Then logError DebugPrint, errors, ("-> " & currentTest & ": dc1(""last"") must be Nothing after re-ordering and insert")
    Next
    
currentTest = "[IIKEC-1] .Item, .Items and .Keys must work with .EmulateCollection=True"
    Set dc1 = New DictCollection: dc1.EmulateCollection = True: dc1.ThrowErrors = False
    Set dc2 = New DictCollection: dc2.DefaultValueEnabled = False ' test object
    dc1.Add "STRING", "ItemThatWillBeOverwritten" ' ext index=1, int index=0
    dc1.Add2 "SecondItem" ' ext_index=2, int_index=1, no key
    dc1.Add "Integer", CInt(1) ' ext_index=3, int_index=2
    dc1.Add "Long", CLng(1) ' ext_index=4, int_index=3
    dc1.Add "Single", CSng(1) ' ext_index=5, int_index=4
    dc1.Add "Double", CDbl(1) ' ext_index=6, int_index=5
    dc1.Add "Currency", CCur(1) ' ext_index=7, int_index=6
    dc1.Add "Decimal", CDec(1) ' ext_index=8, int_index=7
    dc1.Add "Date", CDate(1) ' ext_index=9, int_index=8
    dc1.Add "DictCollection", dc2 ' ext_index=10, int_index=9
    dc1.Add "", "ItemWithEmptyKey" ' ext_index=11, int_index=10, zero length string is supported as key by Collection
    dc1.Add UtilGetMissingValue, "ItemWithNoKey" '_index=12, int_index=11, Missing is not supported as key by Collection
    dc1.Add "String", "FirstItem" ' ext_index=1, int_index=0
    dc1.Item(14) = "LastItem" ' ext_index=14, int_index=13, -> will insert ext_index=13, int_index=12 as Empty/no key
    If dc1.Item(0) <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest & ": accessing invalid indexes with .Item must return '[NONEXISTING]'")
    If dc1.Item(15) <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest & ": accessing invalid indexes with .Item must return '[NONEXISTING]'")
    If dc1.Count <> 14 Then logError DebugPrint, errors, ("-> " & currentTest & ": .Count must be 15")
    If dc1.Item("") <> "ItemWithEmptyKey" Then logError DebugPrint, errors, ("-> " & currentTest & ": zero length string must be allowed as key")
    If dc1.Item(UtilGetMissingValue) <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest & ": Missing must not be allowed as key")
    Items = dc1.Items: Keys = dc1.Keys
        If Items(0) <> "FirstItem" And Keys(0) <> "String" Then logError DebugPrint, errors, ("-> " & currentTest)
        If Items(1) <> "SecondItem" And Keys(1) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
        If TypeName(Items(2)) <> Keys(2) And Items(2) <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
        If TypeName(Items(3)) <> Keys(3) And Items(3) <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
        If TypeName(Items(4)) <> Keys(4) And Items(4) <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
        If TypeName(Items(5)) <> Keys(5) And Items(5) <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
        If TypeName(Items(6)) <> Keys(6) And Items(6) <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
        If TypeName(Items(7)) <> Keys(7) And Items(7) <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
        If TypeName(Items(8)) <> Keys(8) And Items(8) <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
        If TypeName(Items(9)) <> Keys(9) Then
            UtilAddArrayValue errors, ("-> " & currentTest)
        Else
            If Not Items(9) Is dc2 Then logError DebugPrint, errors, ("-> " & currentTest)
        End If
        If Items(10) <> "ItemWithEmptyKey" And Keys(10) <> "" Then logError DebugPrint, errors, ("-> " & currentTest)
        If Items(11) <> "ItemWithNoKey" And Keys(11) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
        If Items(12) <> Empty And Keys(12) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
        If Items(13) <> "LastItem" And Keys(13) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)

currentTest = "[OBI-1] One based index must work with all functions"
    Set dc1 = New DictCollection: dc1.ZeroBasedIndex = False
    dc1(0) = "A": If dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": assigning index=0 must do nothing")
    dc1.Add2 "A", , 0: If dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": adding before 0 must do nothing")
    dc1.Add2 "A", , , -1: If dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": adding after 0 must do nothing")
    dc1.Insert "A", 0: If dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": inserting index=0 must do nothing")
    dc1.Insert "A", 1, "keyA": If dc1.Count <> 1 Then logError DebugPrint, errors, ("-> " & currentTest & ": inserting to index=1 must work") ' dc1.Items = Array("A")
    dc1.Add2 "B", "keyB", , 1: If dc1.Count <> 2 Then logError DebugPrint, errors, ("-> " & currentTest & ": adding after 1 must work") ' dc1.Items = Array("A","B")
    If dc1.Item(1) <> "A" Or dc1.Item(2) <> "B" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.IndexOfKey("keyA") <> 1 Or dc1.IndexOfKey("keyB") <> 2 Then logError DebugPrint, errors, ("-> " & currentTest & ": .IndexOfKey() must return correct index")
    If dc1.KeyOfItemAt(0) <> Empty Or dc1.KeyOfItemAt(3) <> Empty Or dc1.KeyOfItemAt(1) <> "keyA" Or dc1.KeyOfItemAt(2) <> "keyB" Then logError DebugPrint, errors, ("-> " & currentTest & ": .IndexOfKey() must return correct index")
    dc1(4) = "C" ' dc1.Items = Array("A","B",Empty,"C")
    If dc1.Item(1) <> "A" Or dc1.Item(2) <> "B" Or dc1.Item(3) <> Empty Or dc1.Item(4) <> "C" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.ItemAt(1) <> "A" Or dc1.ItemAt(2) <> "B" Or dc1.ItemAt(3) <> Empty Or dc1.ItemAt(4) <> "C" Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Remove 0: If dc1.Count <> 4 Then logError DebugPrint, errors, ("-> " & currentTest & ": removing index=0 must do nothing")
    dc1.Remove 5: If dc1.Count <> 4 Then logError DebugPrint, errors, ("-> " & currentTest & ": removing index=5 must do nothing")
    dc1.Remove 3: If dc1.Count <> 3 Then logError DebugPrint, errors, ("-> " & currentTest & ": removing index=3 must remove the item") ' dc1.Items = Array("A","B","C")
    If dc1.Item(1) <> "A" Or dc1.Item(2) <> "B" Or dc1.Item(3) <> "C" Then logError DebugPrint, errors, ("-> " & currentTest)


currentTest = "[SDC-1] SubDC must copy all settings from parent DictCollection"
    Set dc1 = New DictCollection
    dc1.EmulateDictionary = True
    dc1.ZeroBasedIndex = False
    dc1.CompareMode = vbTextCompare
    dc1.ThrowErrors = False
    dc1.DefaultValueEnabled = False
    dc1.LazySorting = False
    Set dc2 = dc1.subDc("key1")
    If dc2 Is Nothing Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc2.EmulateDictionary <> dc1.EmulateDictionary Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc2.ZeroBasedIndex <> dc1.ZeroBasedIndex Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc2.CompareMode <> dc1.CompareMode Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc2.ThrowErrors <> dc1.ThrowErrors Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc2.DefaultValueEnabled <> dc1.DefaultValueEnabled Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc2.LazySorting <> dc1.LazySorting Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[SDC-2] SubDC must not delete existing DictCollections but overwrite existing non-DictCollection-items"
    Set dc1 = New DictCollection: dc1.Add "keyA", New DictCollection: dc1.Item("keyA").Add "first", 1
    If Not dc1.subDc("keyA").Exists("first") Then logError DebugPrint, errors, ("-> " & currentTest)
    If Not dc1.subDc(0).Exists("first") Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Add "keyB", 123: dc1.subDc("keyB").DefaultValueEnabled = False
    If dc1.Count <> 2 Or TypeName(dc1.Item("keyB")) <> "DictCollection" Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[SDC-3] SubDC must be chainable"
    Set dc1 = New DictCollection
    Err.Clear: dc1.subDc(1).subDc("key1").subDc(0).subDc ("")
    If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.Item(0) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
    If Not IsObject(dc1.Item(1)) Then
        UtilAddArrayValue errors, ("-> " & currentTest)
    Else
        If TypeName(dc1.Item(1)) <> "DictCollection" Then
            UtilAddArrayValue errors, ("-> " & currentTest)
        Else
            ' 4 = filled key-value-store ("key1")
            If dc1.Item(1).CollectionType <> 4 Then logError DebugPrint, errors, ("-> " & currentTest)
            If Not IsObject(dc1.Item(1).Item("key1")) Then
                UtilAddArrayValue errors, ("-> " & currentTest)
            Else
                If TypeName(dc1.Item(1).Item("key1")) <> "DictCollection" Then
                    UtilAddArrayValue errors, ("-> " & currentTest)
                Else
                    ' 2 = filled array
                    If dc1.Item(1).Item("key1").CollectionType <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
                    If Not IsObject(dc1.Item(1).Item("key1").Item(0)) Then
                        UtilAddArrayValue errors, ("-> " & currentTest)
                    Else
                        If TypeName(dc1.Item(1).Item("key1").Item(0)) <> "DictCollection" Then
                            UtilAddArrayValue errors, ("-> " & currentTest)
                        Else
                            ' 2 = filled array
                            If dc1.Item(1).Item("key1").Item(0).CollectionType <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
                            If Not IsObject(dc1.Item(1).Item("key1").Item(0).Item(0)) Then
                                UtilAddArrayValue errors, ("-> " & currentTest)
                            Else
                                If TypeName(dc1.Item(1).Item("key1").Item(0).Item(0)) <> "DictCollection" Then
                                    UtilAddArrayValue errors, ("-> " & currentTest)
                                Else
                                    ' 1 = empty array
                                    If dc1.Item(1).Item("key1").Item(0).Item(0).CollectionType <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    On Error GoTo Fail: Err.Clear
currentTest = "[SDC-4] SubDC must work with .EmulateDictionary=True"
    Set dc1 = New DictCollection: dc1.EmulateDictionary = True: Set dc2 = New DictCollection: dc2.DefaultValueEnabled = False
    testdata = Array(CInt(1), CLng(2), CSng(3.5), CDbl(4.5), CCur(5), CDec(6), "", "x", UtilGetMissingValue, dc2, Null, CVErr(11), Nothing) ' valid Dictionary key types
    For i = 0 To UBound(testdata): dc1.subDc(testdata(i)).Add "", i: Next
    For i = 0 To UBound(testdata)
        If Not dc1.Exists(testdata(i)) Then
            UtilAddArrayValue errors, ("-> " & currentTest)
        Else
            If TypeName(dc1.KeyOfItemAt(i)) <> TypeName(testdata(i)) Then logError DebugPrint, errors, ("-> " & currentTest)
            If Not IsObject(dc1.Item(testdata(i))) Then
                UtilAddArrayValue errors, ("-> " & currentTest)
            Else
                If dc1.Item(testdata(i)).ItemAt(0) <> i Then logError DebugPrint, errors, ("-> " & currentTest)
            End If
        End If
    Next
    On Error Resume Next
    ' try adding existing key
    On Error GoTo Fail: Err.Clear
currentTest = "[SDC-5] SubDC must work with .EmulateCollection=True"
    Set dc1 = New DictCollection: dc1.EmulateCollection = True: Set dc2 = New DictCollection: dc2.DefaultValueEnabled = False
    testdata = Array(Array(1, 2), dc2, Null, CVErr(11), Nothing) ' invalid key types
    On Error Resume Next
    For i = 0 To UBound(testdata)
        Err.Clear: dc1.subDc testdata(i): If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest & ": adding SubCollections with keys of type " & TypeName(testdata(i)) & " must throw Error 13")
    Next
    ' try adding at index = 0 must be possible
    Err.Clear: dc1.subDc 0: If Err.Number <> 9 Or dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": adding SubCollections at index 0 must throw Error 9")
    ' add a DictCollection for a nonexisting key
    Err.Clear: dc1.subDc "": If Err.Number <> 0 Or dc1.Count <> 1 Then logError DebugPrint, errors, ("-> " & currentTest & ": adding SubCollection with empty String as key must work")
    ' access it
    Err.Clear: dc1.subDc "": If Err.Number <> 0 Or dc1.Count <> 1 Then logError DebugPrint, errors, ("-> " & currentTest & ": accessing SubCollection with empty String as key must work")
    ' add another DictCollection for a nonexisting key
    Err.Clear: dc1.subDc "a": If Err.Number <> 0 Or dc1.Count <> 2 Then logError DebugPrint, errors, ("-> " & currentTest & ": adding SubCollection with key must work")
    Keys = dc1.Keys
    If Keys(0) <> "" Or Keys(1) <> "a" Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear


    GoTo ExitFunction
Fail:
    logError DebugPrint, errors, ("-> " & currentTest & " -> Error " & Err.Number & " " & Err.Description)
ExitFunction:
    testAdvancedItemAndKeyAccess = errors
End Function

' Tests advanced functionality (DebugPrint=True -> print errors immediately)
Private Function testAdvancedFunctionality1(Optional DebugPrint As Boolean = False) As Variant
    On Error GoTo Fail: Err.Clear
    Dim errors As Variant, dc1 As DictCollection, dc2 As DictCollection, dc3 As DictCollection, currentTest As String
    Dim i As Long, j As Long, Items As Variant, Keys As Variant, testdata As Variant, val As Variant
    errors = Array()
    
currentTest = "[INC-1] Increment must assign Amount if index or key does not exist"
    Set dc1 = New DictCollection
    dc1.Increment "a", 10: If dc1.Count <> 1 Or dc1.Item("a") <> 10 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Increment 1, -10: If dc1.Count <> 2 Or dc1.Item(1) <> -10 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[INC-2] Increment must add or subtract correct Amount and avoid overflows by implicitly converting types"
    Set dc1 = New DictCollection
    dc1.Item(0) = CInt(30000)
    val = dc1.Increment(0, 30000): If TypeName(dc1.Item(0)) <> "Long" Or dc1.Item(0) <> 60000 Or val <> 60000 Then logError DebugPrint, errors, ("-> " & currentTest)
    val = dc1.Increment(0, 0.5): If TypeName(dc1.Item(0)) <> "Double" Or dc1.Item(0) <> 60000.5 Or val <> 60000.5 Then logError DebugPrint, errors, ("-> " & currentTest)
    val = dc1.Increment(0, -60000.5): If TypeName(dc1.Item(0)) <> "Double" Or dc1.Item(0) <> 0 Or val <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    val = dc1.Increment(0, CDate(15)): If TypeName(dc1.Item(0)) <> "Date" Or dc1.Item(0) <> 15 Or val <> 15 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[INC-3] Increment must throw correct errors"
    Set dc1 = New DictCollection: dc1.ThrowErrors = True
    ' errors are inherited from internalGetItem()
    On Error Resume Next
    ' empty DictCollection
    Err.Clear: val = dc1.Increment(-1): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.Increment(0): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.Increment("a"): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Add "x", 1 ' filled DictCollection
    Err.Clear: val = dc1.Increment(-1): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.Increment("a"): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.Increment(0, Array(1, 2)): If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val = dc1.Increment("x", 1): If Err.Number <> 0 And dc1.Item(0) <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear

currentTest = "[ADCT-1] AnalyzeDictCollectionTree must return correct values"
    Set dc1 = New DictCollection: Set dc2 = dc1.AnalyzeDictCollectionTree()
    ' empty dc
    If dc2("SubItems") <> 0 Or dc2("TreeDepth") <> 0 Or dc2("SubCollections") <> 0 Or dc2("NonStringConvertableItems") <> 0 Or dc2("NonStringConvertableKeys") <> 0 Or dc2("CircularReferences") <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    ' add first dc on level 0
    dc1.AddDC "1": Set dc2 = dc1.AnalyzeDictCollectionTree()
    If dc2("SubItems") <> 1 Or dc2("TreeDepth") <> 1 Or dc2("SubCollections") <> 1 Or dc2("NonStringConvertableItems") <> 0 Or dc2("NonStringConvertableKeys") <> 0 Or dc2("CircularReferences") <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    ' add non-dc item on level 0
    dc1.Add "x", 1: Set dc2 = dc1.AnalyzeDictCollectionTree()
    If dc2("SubItems") <> 2 Or dc2("TreeDepth") <> 1 Or dc2("SubCollections") <> 1 Or dc2("NonStringConvertableItems") <> 0 Or dc2("NonStringConvertableKeys") <> 0 Or dc2("CircularReferences") <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    ' add second dc on level 0
    dc1.AddDC "2": Set dc2 = dc1.AnalyzeDictCollectionTree()
    If dc2("SubItems") <> 3 Or dc2("TreeDepth") <> 1 Or dc2("SubCollections") <> 2 Or dc2("NonStringConvertableItems") <> 0 Or dc2("NonStringConvertableKeys") <> 0 Or dc2("CircularReferences") <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    ' add third dc on level 1
    dc1.AsDC("1").AddDC ("1-1"): Set dc2 = dc1.AnalyzeDictCollectionTree()
    If dc2("SubItems") <> 4 Or dc2("TreeDepth") <> 2 Or dc2("SubCollections") <> 3 Or dc2("NonStringConvertableItems") <> 0 Or dc2("NonStringConvertableKeys") <> 0 Or dc2("CircularReferences") <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    ' add 4th dc on level 1 and 5th dc on level 2
    dc1.AsDC("2").subDc("2-1").AddDC ("2-1-1"): Set dc2 = dc1.AnalyzeDictCollectionTree()
    If dc2("SubItems") <> 6 Or dc2("TreeDepth") <> 3 Or dc2("SubCollections") <> 5 Or dc2("NonStringConvertableItems") <> 0 Or dc2("NonStringConvertableKeys") <> 0 Or dc2("CircularReferences") <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    ' add item without key on level 2
    dc1.AsDC("2").AsDC("2-1").Add , "item": Set dc2 = dc1.AnalyzeDictCollectionTree()
    If dc2("SubItems") <> 7 Or dc2("TreeDepth") <> 3 Or dc2("SubCollections") <> 5 Or dc2("NonStringConvertableItems") <> 0 Or dc2("NonStringConvertableKeys") <> 0 Or dc2("CircularReferences") <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    ' add item without key on level 0
    dc1.AsDC("1").Add , "item": Set dc2 = dc1.AnalyzeDictCollectionTree()
    If dc2("SubItems") <> 8 Or dc2("TreeDepth") <> 3 Or dc2("SubCollections") <> 5 Or dc2("NonStringConvertableItems") <> 0 Or dc2("NonStringConvertableKeys") <> 0 Or dc2("CircularReferences") <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.EmulateDictionary = True
    ' add non string convertable key on level 0
    dc1.Add CVErr(13), "item": Set dc2 = dc1.AnalyzeDictCollectionTree()
    If dc2("SubItems") <> 9 Or dc2("TreeDepth") <> 3 Or dc2("SubCollections") <> 5 Or dc2("NonStringConvertableItems") <> 0 Or dc2("NonStringConvertableKeys") <> 1 Or dc2("CircularReferences") <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    ' add non string convertable key on level 1
    dc1.AsDC("2").EmulateDictionary = True: dc1.AsDC("2").Add New Collection, New Collection:
    Set dc2 = dc1.AnalyzeDictCollectionTree()
    If dc2("SubItems") <> 10 Or dc2("TreeDepth") <> 3 Or dc2("SubCollections") <> 5 Or dc2("NonStringConvertableItems") <> 1 Or dc2("NonStringConvertableKeys") <> 2 Or dc2("CircularReferences") <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    ' add non-circular reference to existing dc on level 0
    dc1.Insert dc1("1"), 0, "c":
    Set dc2 = dc1.AnalyzeDictCollectionTree()
    If dc2("SubItems") <> 13 Or dc2("TreeDepth") <> 3 Or dc2("SubCollections") <> 7 Or dc2("NonStringConvertableItems") <> 1 Or dc2("NonStringConvertableKeys") <> 2 Or dc2("CircularReferences") <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    ' add circular reference on level 0
    dc1.Insert dc1, 0, "d": Set dc2 = dc1.AnalyzeDictCollectionTree()
    If dc2("SubItems") <> 14 Or dc2("TreeDepth") <> 3 Or dc2("SubCollections") <> 8 Or dc2("NonStringConvertableItems") <> 1 Or dc2("NonStringConvertableKeys") <> 2 Or dc2("CircularReferences") <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
    ' add circular reference on level 3
    dc1.AsDC("2").AsDC("2-1").AsDC("2-1-1").Add , dc1
    Set dc2 = dc1.AnalyzeDictCollectionTree()
    If dc2("SubItems") <> 15 Or dc2("TreeDepth") <> 4 Or dc2("SubCollections") <> 9 Or dc2("NonStringConvertableItems") <> 1 Or dc2("NonStringConvertableKeys") <> 2 Or dc2("CircularReferences") <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)

currentTest = "[MOVE-1] Moving items in empty collection must do nothing or throw correct errors"
    On Error Resume Next
    Set dc1 = New DictCollection: dc1.ThrowErrors = False
    Err.Clear: dc1.Move -1, -1: If Err.Number <> 0 Or dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: dc1.Move 0, 1: If Err.Number <> 0 Or dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: dc1.Move 0, 1000: If Err.Number <> 0 Or dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    Set dc1 = New DictCollection: dc1.ThrowErrors = True
    Err.Clear: dc1.Move -1, -1: If Err.Number <> 9 Or dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: dc1.Move 0, 1: If Err.Number <> 9 Or dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: dc1.Move 0, 1000: If Err.Number <> 9 Or dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear
currentTest = "[MOVE-2] Moving items in array"
    Set dc1 = New DictCollection: dc1.ThrowErrors = False
    dc1.Add , 100: dc1.Add , 200: dc1.Add , 300 ' dc1.Items = Array(100,200,300)
    dc1.Move 3, 0 ' must do nothing because FromIndex=3 is out of bounds
    If dc1(0) <> 100 And dc1(1) <> 200 And dc1(2) <> 300 Then logError DebugPrint, errors, ("-> " & currentTest & ": invalid indexes must do nothing")
    dc1.Move 1, -1 ' must do nothing because ToIndex=-1 is out of bounds
    If dc1(0) <> 100 And dc1(1) <> 200 And dc1(2) <> 300 Then logError DebugPrint, errors, ("-> " & currentTest & ": invalid indexes must do nothing")
    dc1.Move 0, 1 ' dc1.Items = Array(200,100,300)
    If dc1(0) <> 200 And dc1(1) <> 100 And dc1(2) <> 300 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Move 0, 2 ' dc1.Items = Array(100,300,200)
    If dc1(0) <> 100 And dc1(1) <> 300 And dc1(2) <> 200 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Move 1, 3 ' dc1.Items = Array(100,200,Empty,300)
    If dc1(0) <> 100 And dc1(1) <> 200 And Not IsEmpty(dc1(2)) And dc1(3) <> 300 Then logError DebugPrint, errors, ("-> " & currentTest)
    Items = dc1.Items
    If Items(0) <> 100 And Items(1) <> 200 And Not IsEmpty(Items(2)) And Items(3) <> 300 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[MOVE-3] Moving items in key-value-store"
    Set dc1 = New DictCollection: dc1.ThrowErrors = False
    dc1.Add "A", 100: dc1.Add "B", 200: dc1.Add "C", 300 ' dc1.Items = Array(100,200,300) dc1.Keys("A","B","C")
    dc1.Move 0, 2 ' dc1.Items = Array(200,300,100) dc1.Keys("B","C","A")
    Keys = dc1.Keys
    If dc1(0) <> 200 And dc1(1) <> 300 And dc1(2) <> 100 Then logError DebugPrint, errors, ("-> " & currentTest)
    If Keys(0) <> "B" And Keys(1) <> "C" And Keys(2) <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Move 1, 0 ' dc1.Items = Array(300,200,100) dc1.Keys("C","B","A")
    Keys = dc1.Keys
    If dc1(0) <> 300 And dc1(1) <> 200 And dc1(2) <> 100 Then logError DebugPrint, errors, ("-> " & currentTest)
    If Keys(0) <> "C" And Keys(1) <> "B" And Keys(2) <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Move 0, 3 ' dc1.Items = Array(200,100,Empty,300) dc1.Keys("B","A",Empty,"C")
    Keys = dc1.Keys
    If dc1(0) <> 200 And dc1(1) <> 100 And Not IsEmpty(dc1(2)) And dc1(3) <> 300 Then logError DebugPrint, errors, ("-> " & currentTest)
    If Keys(0) <> "B" And Keys(1) <> "A" And Not IsEmpty(Keys(2)) And Keys(3) <> "C" Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[MOVE-4] Moving items must work with .EmulateDictionary=True"
    Set dc1 = New DictCollection: dc1.EmulateDictionary = True: dc1.ThrowErrors = False
    Set dc2 = New DictCollection: dc2.DefaultValueEnabled = False ' object key
    dc1.Add CInt(100), "A": dc1.Add CDate(200), "B": dc1.Add dc2, "C" ' dc1.Items = Array("A","B","C") dc1.Keys(CInt(100),CDate(200),dc2)
    dc1.Move 0, 2 ' dc1.Items = Array("B","C","A") dc1.Keys(CDate(200),dc2,CInt(100))
    Keys = dc1.Keys
    If dc1(CDate(200)) <> "B" Or dc1(dc2) <> "C" Or dc1(CInt(100)) <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.ItemAt(0) <> "B" Or dc1.ItemAt(1) <> "C" Or dc1.ItemAt(2) <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)
    If TypeName(Keys(0)) <> "Date" Or TypeName(Keys(1)) <> "DictCollection" Or TypeName(Keys(2)) <> "Integer" Then
        UtilAddArrayValue errors, ("-> " & currentTest)
    Else
        If Not Keys(1) Is dc2 Then logError DebugPrint, errors, ("-> " & currentTest)
    End If
    dc1.Move 2, 3 ' dc1.Items = Array("B","C",Empty,"A") dc1.Keys(CDate(200),dc2,Empty,CInt(100))
    Keys = dc1.Keys
    If dc1(CDate(200)) <> "B" Or dc1(dc2) <> "C" Or dc1(CInt(100)) <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.ItemAt(0) <> "B" Or dc1.ItemAt(1) <> "C" Or Not IsEmpty(dc1.ItemAt(2)) Or dc1.ItemAt(3) <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)
    If TypeName(Keys(0)) <> "Date" Or TypeName(Keys(1)) <> "DictCollection" Or TypeName(Keys(2)) <> "Empty" Or TypeName(Keys(3)) <> "Integer" Then
        UtilAddArrayValue errors, ("-> " & currentTest)
    Else
        If Not Keys(1) Is dc2 Then logError DebugPrint, errors, ("-> " & currentTest)
    End If
currentTest = "[MOVE-5] Moving items must work with .EmulateCollection=True"
    Set dc1 = New DictCollection: dc1.EmulateCollection = True: dc1.ThrowErrors = False
    dc1.Add2 100, "A": dc1.Add2 200, "B": dc1.Add2 300, "": dc1.Add2 400, Empty: dc1.Add2 500 ' dc1.Items = Array(100,200,300,400,500) dc1.Keys("A","B","",Empty,Empty)
    dc1.Move 0, 2 ' must do nothing because FromIndex=0 is out of bounds
    If dc1(1) <> 100 And dc1(2) <> 200 And dc1(3) <> 300 And dc1(4) <> 400 And dc1(5) <> 500 Then logError DebugPrint, errors, ("-> " & currentTest & ": invalid indexes must do nothing")
    dc1.Move 2, 0 ' must do nothing because ToIndex=0 is out of bounds
    If dc1(1) <> 100 And dc1(2) <> 200 And dc1(3) <> 300 And dc1(4) <> 400 And dc1(5) <> 500 Then logError DebugPrint, errors, ("-> " & currentTest & ": invalid indexes must do nothing")
    dc1.Move 1, 4 ' dc1.Items = Array(200,300,400,100,500) dc1.Keys("B","",Empty,"A",Empty)
    Keys = dc1.Keys
    If dc1(1) <> 200 And dc1(2) <> 300 And dc1(3) <> 400 And dc1(4) <> 100 And dc1(5) <> 500 Then logError DebugPrint, errors, ("-> " & currentTest)
    If Keys(0) <> "B" Or Keys(1) <> "" Or TypeName(Keys(1)) <> "String" Or Not IsEmpty(Keys(2)) Or Keys(3) <> "A" Or Not IsEmpty(Keys(4)) Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Move 5, 1 ' dc1.Items = Array(500,200,300,400,100) dc1.Keys(Empty,"B","",Empty,"A")
    Keys = dc1.Keys: Items = dc1.Items
    If dc1(1) <> 500 And dc1(2) <> 200 And dc1(3) <> 300 And dc1(4) <> 400 And dc1(5) <> 100 Then logError DebugPrint, errors, ("-> " & currentTest)
    If Not IsEmpty(Keys(0)) Or Keys(1) <> "B" Or Keys(2) <> "" Or TypeName(Keys(2)) <> "String" Or Not IsEmpty(Keys(3)) Or Keys(4) <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Move 2, 6 ' dc1.Items = Array(500,300,400,100,Empty,200) dc1.Keys(Empty,"",Empty,"A",Empty,"B")
    Keys = dc1.Keys
    If dc1(1) <> 500 And dc1(2) <> 300 And dc1(3) <> 400 And dc1(4) <> 100 And Not IsEmpty(dc1(5)) And dc1(6) <> 200 Then logError DebugPrint, errors, ("-> " & currentTest)
    If Not IsEmpty(Keys(0)) Or TypeName(Keys(1)) <> "String" Or Keys(1) <> "" Or Not IsEmpty(Keys(2)) Or Keys(3) <> "A" Or Not IsEmpty(Keys(4)) Or Keys(5) <> "B" Then logError DebugPrint, errors, ("-> " & currentTest)

    ' the expression '(Not Not TestArray) <> 0' evaluates as True if an array is initialized and as False if it is uninitialized; works only with arrays declared as Type() (explicit array)
currentTest = "[MCT-1] Modifying Collection Type: Change to nonexisting must delete all items and keys"
    Set dc1 = New DictCollection: dc1.Add , "itemA": dc1.CollectionType = 0  ' set to nonexisting
    Items = dc1.Items: Keys = dc1.Keys: If dc1.Count <> 0 Or dc1.KeyCount <> 0 Or UBound(Items) <> -1 Or UBound(Keys) <> -1 Then logError DebugPrint, errors, ("-> " & currentTest)
    'intItems = dc1.InternalItems: intKeys = dc1.InternalKeys: If (Not Not intItems) <> 0 Or (Not Not intKeys) <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    Set dc1 = New DictCollection: dc1.Add "keyA", "itemA": dc1.CollectionType = 0  ' set to nonexisting
    Items = dc1.Items: Keys = dc1.Keys: If dc1.Count <> 0 Or dc1.KeyCount <> 0 Or UBound(Items) <> -1 Or UBound(Keys) <> -1 Then logError DebugPrint, errors, ("-> " & currentTest)
    'intItems = dc1.InternalItems: intKeys = dc1.InternalKeys: If (Not Not intItems) <> 0 Or (Not Not intKeys) <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[MCT-2] Modifying Collection Type: Change from filled array to empty array must delete all items"
    Set dc1 = New DictCollection: dc1.Add , "itemA"
    If dc1.CollectionType <> 2 Then logError DebugPrint, errors, ("-> " & currentTest & ": .CollectionType must equal 2")
    dc1.CollectionType = 1  ' set empty array
    Items = dc1.Items: Keys = dc1.Keys: If dc1.Count <> 0 Or dc1.KeyCount <> 0 Or UBound(Items) <> -1 Or UBound(Keys) <> -1 Then logError DebugPrint, errors, ("-> " & currentTest)
    'intItems = dc1.InternalItems: intKeys = dc1.InternalKeys: If (Not Not intItems) = 0 Or (Not Not intKeys) <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[MCT-3] Modifying Collection Type: Change from filled key-value-store to empty key-value-store must delete all items and keys"
    Set dc1 = New DictCollection: dc1.Add "keyA", "itemA"
    If dc1.CollectionType <> 4 Then logError DebugPrint, errors, ("-> " & currentTest & ": .CollectionType must equal 4")
    dc1.CollectionType = 3  ' set to empty key-value-store
    Items = dc1.Items: Keys = dc1.Keys: If dc1.Count <> 0 Or dc1.KeyCount <> 0 Or UBound(Items) <> -1 Or UBound(Keys) <> -1 Then logError DebugPrint, errors, ("-> " & currentTest)
    'intItems = dc1.InternalItems: intKeys = dc1.InternalKeys: If (Not Not intItems) = 0 Or (Not Not intKeys) = 0 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[MCT-4] Modifying Collection Type: Change from filled key-value-store to array must delete all keys"
    Set dc1 = New DictCollection: dc1.Add "keyA", "itemA"
    dc1.CollectionType = 2  ' set to filled array
    Items = dc1.Items: Keys = dc1.Keys: If dc1.Count <> 1 Or dc1.KeyCount <> 0 Or UBound(Items) <> 0 Or UBound(Keys) <> -1 Then logError DebugPrint, errors, ("-> " & currentTest)
    'intItems = dc1.InternalItems: intKeys = dc1.InternalKeys: If (Not Not intItems) = 0 Or (Not Not intKeys) <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[MCT-5] Modifying Collection Type: Change from filled key-value-store with items having no key to from filled key-value-store must add keys"
    Set dc1 = New DictCollection: dc1.Add , "itemA": dc1.Add "keyB", "itemB"
    If dc1.CollectionType <> 5 Then logError DebugPrint, errors, ("-> " & currentTest & ": .CollectionType must equal 5")
    dc1.CollectionType = 4  ' set to filled key-value-store
    If dc1.CollectionType <> 4 Then logError DebugPrint, errors, ("-> " & currentTest & ": .CollectionType must equal 4")
    Items = dc1.Items: Keys = dc1.Keys: If dc1.Count <> 2 Or dc1.KeyCount <> 2 Or UBound(Items) <> 1 Or UBound(Keys) <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
    'intItems = dc1.InternalItems: intKeys = dc1.InternalKeys: If (Not Not intItems) = 0 Or (Not Not intKeys) = 0 Then logError DebugPrint, errors, ("-> " & currentTest)

For i = 1 To 2
    If i = 1 Then
currentTest = "[GIK-1] Generating Item Keys: Keys must be generated from index as '_POS'"
    Else
currentTest = "[EIHK-1] EnsureAllItemsHaveKeys: Keys must be generated from index as '_POS'"
    End If
    Set dc1 = New DictCollection: dc1.Add , "A": dc1.Add "keyB", "B": dc1.Add , "C": dc1.Item(4) = "D" ' dc1.Items = Array("A","B","C",Empty,"D") dc1.Keys = Array(Empty,"keyB",Empty,Empty,Empty)
    If i = 1 Then
        dc1.CollectionType = 4 ' set to filled key-value-store
    Else
        dc1.EnsureAllItemsHaveKeys
    End If
    If dc1.Count <> 5 Or dc1.KeyCount <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.Item("_1") <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.Item("_2") <> dc1.NonExistingValue Or dc1.Item("keyB") <> "B" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.Item("_3") <> "C" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.Item("_4") <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.Item("_5") <> "D" Then logError DebugPrint, errors, ("-> " & currentTest)
    If i = 1 Then
currentTest = "[GIK-2] Generating Item Keys: Existing Keys in format '_IDX' and '_IDX_N' must not be overwritten"
    Else
currentTest = "[EIHK-2] EnsureAllItemsHaveKeys: Existing Keys in format '_IDX' and '_IDX_N' must not be overwritten"
    End If
    Set dc1 = New DictCollection: dc1.Add , "A": dc1.Add "_1", "B": dc1.Add "_1_1", "C": dc1.Add "_1_2", "D" ' dc1.Items = Array("A","B","C","D") dc1.Keys = Array(Empty,"_1","_1_1","_1_2")
    If i = 1 Then
        dc1.CollectionType = 4 ' set to filled key-value-store
    Else
        dc1.EnsureAllItemsHaveKeys
    End If
    If dc1.Item("_1_3") <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.Item("_1") <> "B" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.Item("_1_1") <> "C" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.Item("_1_2") <> "D" Then logError DebugPrint, errors, ("-> " & currentTest)
Next

    
currentTest = "[DPE-1] Accessing Default Property must throw error when using invalid key types"
    Set dc1 = New DictCollection: dc1.DefaultValueEnabled = True: dc1.ThrowErrors = True
    On Error Resume Next:
    Err.Clear: dc1(Null) = "A": If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: dc1(CVErr(5)) = "A": If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest)
    
currentTest = "[DVS-1] Enabling default value must prevent usage of object expressions and allow setting/getting items using default property"
    Set dc1 = New DictCollection: dc1.DefaultValueEnabled = True: dc1.Add "", "A"
    If dc1.Item(UtilGetMissingValue) <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    ' TypeName(dc1) returns "DictCollection" instead of "String" because of Default Property evaluation
    If TypeName(dc1) <> "DictCollection" Then logError DebugPrint, errors, ("-> " & currentTest)
    If VarType(dc1) <> vbString Or dc1 <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)
    ' using default property to set default value
    dc1 = CInt(1): If VarType(dc1) <> vbInteger Or dc1 <> 1 Or dc1.ItemAt(0) <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1() = CDbl(-1): If VarType(dc1()) <> vbDouble Or dc1() <> -1 Or dc1.ItemAt(0) <> -1 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1(UtilGetMissingValue) = CDate(36526): If VarType(dc1(UtilGetMissingValue)) <> vbDate Or dc1(UtilGetMissingValue) <> 36526 Or dc1.ItemAt(0) <> 36526 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1(Null) = "B": If dc1.Count <> 1 Then logError DebugPrint, errors, ("-> " & currentTest & ": using invalid keys with default property must do nothing")
    dc1(1) = "C": If dc1.Count <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1(0) = "D": If dc1.Count <> 3 And VarType(dc1) <> vbString Or dc1 <> "D" Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error Resume Next: Err.Clear
    ' test if setting an object reference works
    Set dc2 = dc1: If Err.Number <> 0 Or Not dc2 Is dc1 Then logError DebugPrint, errors, ("-> " & currentTest & ": Set expression must work with DictCollection")
    If Not dc2 Is dc1 Then logError DebugPrint, errors, ("-> " & currentTest)
    ' test if assigning dc2 as the first item (default value) of dc1 works
    Set dc2 = New DictCollection
    Err.Clear: dc1 = dc2 ' assign using default property (without Set keyword)
    If TypeName(dc1()) <> "DictCollection" Or Err.Number <> 0 Then
        UtilAddArrayValue errors, ("-> " & currentTest)
    Else
        If Not dc1() Is dc2 Or Not dc1.ItemAt(0) Is dc2 Then logError DebugPrint, errors, ("-> " & currentTest)
    End If
    dc1.RemoveAll
    Err.Clear: Set dc2 = dc1.Item(UtilGetMissingValue): If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": setting other DictCollections to [NONEXISTING] (DictCollections) must work")
    Err.Clear: Set dc2 = dc1.Item(Null): If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": setting other DictCollections to [NONEXISTING] (DictCollections) must work")
    Err.Clear: Set dc2 = dc1.Item(Array(1, 2)): If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": setting other DictCollections to [NONEXISTING] (DictCollections) must work")
    Err.Clear: Set dc2 = dc1.Item("nonexisting"): If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": setting other DictCollections to [NONEXISTING] (DictCollections) must work")
    Err.Clear: Set dc2 = dc1(UtilGetMissingValue): If Err.Number <> 424 Then logError DebugPrint, errors, ("-> " & currentTest & ": assigning objects with dc() must throw error 424 'Object required'")
    Err.Clear: Set dc2 = dc1(): If Err.Number <> 424 Then logError DebugPrint, errors, ("-> " & currentTest & ": assigning objects with dc() must throw error 424 'Object required'")
    On Error GoTo Fail
currentTest = "[DVS-2] Disabling default value must allow use of object expressions and prevent setting/getting default values using default property"
    Set dc1 = New DictCollection: dc1.DefaultValueEnabled = False: dc1.Add "", "A"
    If TypeName(dc1) <> "DictCollection" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.Item(UtilGetMissingValue) <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    If TypeName(dc1(UtilGetMissingValue)) <> "DictCollection" Then logError DebugPrint, errors, ("-> " & currentTest)
    If TypeName(dc1()) <> "DictCollection" Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error Resume Next
    Set dc2 = New DictCollection: Set dc2 = dc1: If Err.Number <> 0 Or Not (dc2 Is dc1) Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: Set dc2 = New DictCollection: Set dc2 = dc1.Item(UtilGetMissingValue): If Err.Number <> 424 Or dc2 Is dc1 Then logError DebugPrint, errors, ("-> " & currentTest & ": assigning objects with dc() must throw error 424 'Object required'")
    Err.Clear: Set dc2 = New DictCollection: Set dc2 = dc1(UtilGetMissingValue): If Err.Number <> 0 Or Not (dc2 Is dc1) Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: Set dc2 = New DictCollection: Set dc2 = dc1(): If Err.Number <> 0 Or Not (dc2 Is dc1) Then logError DebugPrint, errors, ("-> " & currentTest)
    ' setting default value using default property must fail
    Set dc1 = New DictCollection: dc1.DefaultValueEnabled = False
    dc1 = CInt(1): If dc1.CollectionType <> 1 And dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1() = CDbl(-1): If dc1 <> dc1.NonExistingValue And dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1(UtilGetMissingValue) = CDate(36526): If dc1 <> dc1.NonExistingValue And dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1(Null) = "A": If dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": using invalid keys with default property must do nothing")
    dc1(1) = "B": If dc1.Count <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1(0) = "C": If dc1.Count <> 2 And TypeName(dc1) <> "DictCollection" Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear
currentTest = "[DVS-3] Disabling default value must throw errors when attemting to set default values using default property"
    On Error Resume Next
    Set dc1 = New DictCollection: dc1.DefaultValueEnabled = False: dc1.ThrowErrors = True
    Err.Clear: dc1 = CInt(1): If Err.Number <> 5 Or dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: dc1() = CDbl(-1): If Err.Number <> 5 Or dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: dc1(UtilGetMissingValue) = CDate(1): If Err.Number <> 5 Or dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear
currentTest = "[DVS-4] Disabling default value must return [NONEXISTING] when accessing nonexisting items and prevent subcollection chaining"
    Set dc1 = New DictCollection: dc1.DefaultValueEnabled = False
    On Error Resume Next
    Err.Clear: If dc1("nonexisting") <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: testdata = dc1("B")("C"): If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear
    
currentTest = "[CSCS-1] CopySubCollectionChainingSettings must copy correct settings"
    Set dc1 = New DictCollection: dc1.ThrowErrors = True: dc1.EmptyCollectionValue = True: dc1.NonExistingValue = CDbl(1.5)
    Set dc2 = New DictCollection: dc2.EmulateDictionary = True: dc2.ZeroBasedIndex = False: dc2.LazySorting = False
    dc1.CopySubCollectionChainingSettings dc1, dc2
    If TypeName(dc2.EmptyCollectionValue) <> "Boolean" Or dc2.EmptyCollectionValue <> True Then logError DebugPrint, errors, ("-> " & currentTest)
    If TypeName(dc2.NonExistingValue) <> "Double" Or dc2.NonExistingValue <> 1.5 Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc2.ThrowErrors <> True Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc2.EmulateDictionary <> True Or dc2.ZeroBasedIndex <> False Or dc2.LazySorting <> False Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[CAS-1] Copying all settings must work"
    Set dc1 = New DictCollection:
    dc1.ThrowErrors = Not THROW_ERRORS_DEFAULT
    dc1.DefaultValueEnabled = Not DEFAULTVALUEENABLED_DEFAULT
    dc1.ZeroBasedIndex = Not ZEROBASEDINDEX_DEFAULT
    dc1.LazySorting = Not LAZYSORTING_DEFAULT
    dc1.EmptyCollectionValue = CDate(1)
    dc1.NonExistingValue = Null
    dc1.CompareMode = 2 ' database
    Set dc2 = New DictCollection: dc1.CopyAllSettingsExceptCollectionType dc1, dc2
        If dc2.ThrowErrors = THROW_ERRORS_DEFAULT Or dc2.DefaultValueEnabled = DEFAULTVALUEENABLED_DEFAULT Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc2.ZeroBasedIndex = ZEROBASEDINDEX_DEFAULT Or dc2.LazySorting = LAZYSORTING_DEFAULT Then logError DebugPrint, errors, ("-> " & currentTest)
        If TypeName(dc2.EmptyCollectionValue) <> "Date" Or dc2.EmptyCollectionValue <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
        If TypeName(dc2.NonExistingValue) <> "Null" Or dc2.NonExistingValue <> Null Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc2.CompareMode <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc2.EmulateCollection <> False Or dc2.EmulateDictionary <> False Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[CAS-2] Copying all settings must work when modifying settings after setting .EmulateCollection=True"
    dc1.EmulateCollection = True
    dc1.ThrowErrors = False
    dc1.DefaultValueEnabled = True
    dc1.ZeroBasedIndex = False
    dc1.CompareMode = 2 ' database
    Set dc2 = New DictCollection: dc1.CopyAllSettingsExceptCollectionType dc1, dc2
        If dc2.ThrowErrors <> False Or dc2.DefaultValueEnabled <> True Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc2.ZeroBasedIndex <> False Then logError DebugPrint, errors, ("-> " & currentTest)
        If TypeName(dc2.EmptyCollectionValue) <> "Date" Or dc2.EmptyCollectionValue <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
        If TypeName(dc2.NonExistingValue) <> "Null" Or dc2.NonExistingValue <> Null Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc2.CompareMode <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[CAS-3] Copying all settings must work when modifying settings after setting .EmulateDictionary=True"
    dc1.EmulateDictionary = True
    dc1.ThrowErrors = False
    dc1.DefaultValueEnabled = True
    dc1.ZeroBasedIndex = True
    dc1.EmptyCollectionValue = "a"
    dc1.NonExistingValue = "b"
    dc1.CompareMode = 2 ' database
    Set dc2 = New DictCollection: dc1.CopyAllSettingsExceptCollectionType dc1, dc2
        If dc2.ThrowErrors <> False Or dc2.DefaultValueEnabled <> True Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc2.ZeroBasedIndex <> True Then logError DebugPrint, errors, ("-> " & currentTest)
        If TypeName(dc2.EmptyCollectionValue) <> "String" Or dc2.EmptyCollectionValue <> "a" Then logError DebugPrint, errors, ("-> " & currentTest)
        If TypeName(dc2.NonExistingValue) <> "String" Or dc2.NonExistingValue <> "b" Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc2.CompareMode <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    

currentTest = "[INSERT-1] Inserting items in empty or nonexisting collections/arrays"
    Set dc1 = New DictCollection: dc1.CollectionType = 0 ' nonexisting
    dc1.Insert: If dc1.Count <> 1 Or dc1(0) <> Empty Or dc1.CollectionType <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.CollectionType = 1 ' empty array
    dc1.Insert: If dc1.Count <> 1 Or dc1(0) <> Empty Or dc1.CollectionType <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.RemoveAll
    dc1.Insert "A", 1, "keyA" ' dc1.Items = Array(Empty,"A") dc1.Keys = Array(Empty,"keyA")
        If dc1.Count <> 2 Or dc1(0) <> Empty Or dc1(1) <> "A" Or dc1("keyA") <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)
        Items = dc1.Items: Keys = dc1.Keys
        If Items(0) <> Empty Or Items(1) <> "A" Or Keys(0) <> Empty Or Keys(1) <> "keyA" Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Insert "B", 1, "keyB" ' dc1.Items = Array(Empty,"B","A") dc1.Keys = Array(Empty,"keyB","keyA")
        If dc1.Count <> 3 Or dc1(0) <> Empty Or dc1(1) <> "B" Or dc1("keyB") <> "B" Or dc1(2) <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)
        Items = dc1.Items: Keys = dc1.Keys
        If Items(0) <> Empty Or Items(1) <> "B" Or Items(2) <> "A" Then logError DebugPrint, errors, ("-> " & currentTest)
        If Keys(0) <> Empty Or Keys(1) <> "keyB" Or Keys(2) <> "keyA" Then logError DebugPrint, errors, ("-> " & currentTest)

currentTest = "[INSERT-2] Inserting items with existing key must remove the item that was previously accociated with that key"
    dc1.CompareMode = 1 ' case insensitive text compare to test correct key overwrite
    dc1.Insert "C", 1, "KEYA" ' dc1.Items = Array(Empty,"C","B") dc1.Keys = Array(Empty,"KEYA","keyB")
        If dc1.Count <> 3 Or dc1(0) <> Empty Or dc1(1) <> "C" Or dc1("KEYA") <> "C" Or dc1(2) <> "B" Then logError DebugPrint, errors, ("-> " & currentTest)
        Items = dc1.Items: Keys = dc1.Keys
        If Items(0) <> Empty Or Items(1) <> "C" Or Items(2) <> "B" Then logError DebugPrint, errors, ("-> " & currentTest)
        If Keys(0) <> Empty Or Keys(1) <> "KEYA" Or Keys(2) <> "keyB" Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Insert "D", 3, "keya" ' dc1.Items = Array(Empty,"B",Empty,"D") dc1.Keys = Array(Empty,"keyB",Empty,"keya")
        If dc1.Count <> 4 Or dc1(0) <> Empty Or dc1(1) <> "B" Or dc1("keya") <> "D" Or dc1(2) <> Empty Or dc1(3) <> "D" Then logError DebugPrint, errors, ("-> " & currentTest)
        Items = dc1.Items: Keys = dc1.Keys
        If Items(0) <> Empty Or Items(1) <> "B" Or Items(2) <> Empty Or Items(3) <> "D" Then logError DebugPrint, errors, ("-> " & currentTest)
        If Keys(0) <> Empty Or Keys(1) <> "keyB" Or Keys(2) <> Empty Or Keys(3) <> "keya" Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Insert "E", 0, "keyE" ' dc1.Items = Array("E",Empty,"B",Empty,"D") dc1.Keys = Array("keyE",Empty,"keyB",Empty,"keya")
        If dc1.Count <> 5 Or dc1(0) <> "E" Or dc1(1) <> Empty Or dc1(2) <> "B" Or dc1("keya") <> "D" Or dc1(3) <> Empty Or dc1(4) <> "D" Then logError DebugPrint, errors, ("-> " & currentTest)
        Items = dc1.Items: Keys = dc1.Keys
        If Items(0) <> "E" Or Items(1) <> Empty Or Items(2) <> "B" Or Items(3) <> Empty Or Items(4) <> "D" Then logError DebugPrint, errors, ("-> " & currentTest)
        If Keys(0) <> "keyE" Or Keys(1) <> Empty Or Keys(2) <> "keyB" Or Keys(3) <> Empty Or Keys(4) <> "keya" Then logError DebugPrint, errors, ("-> " & currentTest)

currentTest = "[INSERT-3] Inserting items must work with .EmulateDictionary=True"
    Set dc1 = New DictCollection: dc1.EmulateDictionary = True: Set dc2 = New DictCollection: dc2.DefaultValueEnabled = False
    On Error Resume Next
    ' test if inserting duplicate keys throws correct errors
    Err.Clear: dc1.RemoveAll: dc1.Insert , , CInt(1): dc1.Insert , , CDbl(1)
    If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & ": inserting a numeric key twice must throw Error 457 'This key is already associated with an element of this collection'")
    Err.Clear: dc1.RemoveAll: dc1.Insert , , "x": dc1.Insert , , "x"
    If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & ": inserting a String key twice must throw Error 457 'This key is already associated with an element of this collection'")
    Err.Clear: dc1.RemoveAll
    dc1.Insert: If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest) ' dc1.Items = Array(Empty) dc1.Keys = Array(Missing)
        If dc1.Count <> 1 Or dc1.ItemAt(0) <> Empty Or Not IsMissing(dc1.KeyOfItemAt(0)) Or dc1(UtilGetMissingValue) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest & ": inserting 1 item with Missing as key must work")
    Err.Clear: dc1.Insert , , UtilGetMissingValue: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & ": inserting a second item with Missing as key must throw Error 457 'This key is already associated with an element of this collection'")
    Err.Clear: dc1.Insert "A", -2: If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest & ": inserting an item at -2 must throw Error 9 'Subscript out of range'")
    On Error GoTo Fail
    ' test if inserting items works like intended
    dc1.Insert "A", 0, "keyA" ' dc1.Items = Array("A",Empty) dc1.Keys = Array("keyA",Missing)
        If dc1.Count <> 2 Or dc1.ItemAt(0) <> "A" Or dc1.KeyOfItemAt(0) <> "keyA" Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc1.ItemAt(1) <> Empty Or Not IsMissing(dc1.KeyOfItemAt(1)) Then logError DebugPrint, errors, ("-> " & currentTest)
        Items = dc1.Items: Keys = dc1.Keys
        If Items(0) <> "A" Or Items(1) <> Empty Or Keys(0) <> "keyA" Or Not IsMissing(Keys(1)) Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Insert "B", 3, "keya" ' dc1.Items = Array("A",Empty,Empty,"B") dc1.Keys = Array("keyA",Missing,Empty,"keya")
        If dc1.Count <> 4 Or dc1.ItemAt(0) <> "A" Or dc1.ItemAt(1) <> Empty Or dc1.ItemAt(2) <> Empty Or dc1.ItemAt(3) <> "B" Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc1.KeyOfItemAt(0) <> "keyA" Or Not IsMissing(dc1.KeyOfItemAt(1)) Or dc1.KeyOfItemAt(2) <> Empty Or dc1.KeyOfItemAt(3) <> "keya" Then logError DebugPrint, errors, ("-> " & currentTest)
    ' test if key data types work and are stored/retrieved correctly
    dc1.RemoveAll: testdata = Array(CInt(1), CLng(2), CSng(3), CDbl(4), CCur(5), CDec(6), CDate(7), Null, CVErr(11), False, dc2, Nothing)
    For i = 0 To UBound(testdata): dc1.Insert i, , testdata(i): Next ' item = i
    If dc1.Count <> UBound(testdata) + 1 Or dc1.KeyCount <> UBound(testdata) + 1 Then logError DebugPrint, errors, ("-> " & currentTest)
    For i = 0 To UBound(testdata)
        If TypeName(dc1.KeyOfItemAt(i)) <> TypeName(testdata(i)) Then logError DebugPrint, errors, ("-> " & currentTest)
        If TypeName(dc1.KeyOfItemAt(i)) = "DictCollection" Then
            If Not dc1.KeyOfItemAt(i) Is dc2 Then logError DebugPrint, errors, ("-> " & currentTest)
        End If
        If dc1(testdata(i)) <> i Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc1.IndexOfKey(testdata(i)) <> i Then logError DebugPrint, errors, ("-> " & currentTest)
    Next

currentTest = "[INSERT-4] Inserting items must work with .EmulateCollection=True"
    Set dc1 = New DictCollection: dc1.EmulateCollection = True: Set dc2 = New DictCollection: dc2.DefaultValueEnabled = False
    On Error Resume Next
    ' test if inserting duplicate keys throws correct errors
    Err.Clear: dc1.RemoveAll: dc1.Insert , , "": dc1.Insert , , ""
    If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & ": inserting a key twice must throw Error 457 'This key is already associated with an element of this collection'")
    Err.Clear: dc1.RemoveAll: dc1.Insert , , "a": dc1.Insert , , "a"
    If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & ": inserting a key twice must throw Error 457 'This key is already associated with an element of this collection'")
    ' test if wrong key data type throws correct error and if key = Missing works as intended
    dc1.RemoveAll: testdata = Array(CInt(1), CLng(2), CSng(3), CDbl(4), CCur(5), CDec(6), CDate(7), Null, CVErr(11), False, dc2, Nothing)
    For i = 0 To UBound(testdata)
        Err.Clear: dc1.Insert , , testdata(i): If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest & ": inserting key type " & TypeName(testdata(i)) & " must throw Error 13 'Type mismatch'")
    Next
    Err.Clear: dc1.Insert "a", , UtilGetMissingValue: If Err.Number <> 0 Or dc1.Count <> 1 Or dc1.KeyCount <> 0 Or dc1(1) <> "a" Then logError DebugPrint, errors, ("-> " & currentTest & ": using Missing as key (no key) must work")
    ' test if using invalid index throws correct errors
    Err.Clear: dc1.Insert "a", -2: If dc2.Count <> 0 Or Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest & ": using -2 as index must throw Error 9 'Subscript out of range'")
    Err.Clear: dc1.Insert "a", 0: If dc2.Count <> 0 Or Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest & ": using -0 as index must throw Error 9 'Subscript out of range'")
    On Error GoTo Fail
    ' test if inserting items works like intended
    dc1.RemoveAll: dc1.Insert: dc1.Insert "A", 1, "keyA" ' dc1.Items = Array("A",Empty) dc1.Keys = Array("keyA",Empty)
        If dc1.Count <> 2 Or dc1(1) <> "A" Or dc1.ItemAt(1) <> "A" Or dc1.KeyOfItemAt(1) <> "keyA" Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc1(2) <> Empty Or dc1.ItemAt(2) <> Empty Or dc1.KeyOfItemAt(2) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
        Items = dc1.Items: Keys = dc1.Keys
        If Items(0) <> "A" Or Items(1) <> Empty Or Keys(0) <> "keyA" Or Keys(1) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.Insert "B", 4, "" ' dc1.Items = Array("A",Empty,Empty,"B") dc1.Keys = Array("keyA",Empty,Empty,"")
        If dc1.Count <> 4 Or dc1(1) <> "A" Or dc1(2) <> Empty Or dc1(3) <> Empty Or dc1(4) <> "B" Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc1.KeyOfItemAt(1) <> "keyA" Or dc1.KeyOfItemAt(2) <> Empty Or dc1.KeyOfItemAt(3) <> Empty Or TypeName(dc1.KeyOfItemAt(4)) <> "String" Or dc1.KeyOfItemAt(4) <> "" Then logError DebugPrint, errors, ("-> " & currentTest)
        Items = dc1.Items: Keys = dc1.Keys
        If Items(0) <> "A" Or Items(1) <> Empty Or Items(2) <> Empty Or Items(3) <> "B" Then logError DebugPrint, errors, ("-> " & currentTest)
        If Keys(0) <> "keyA" Or Keys(1) <> Empty Or Keys(2) <> Empty Or TypeName(Keys(3)) <> "String" Or Keys(3) <> "" Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc1.IndexOfKey("keyA") <> 1 Or dc1.IndexOfKey("") <> 4 Then logError DebugPrint, errors, ("-> " & currentTest)

currentTest = "[ADDDC-1] AddDC must copy all settings from parent DictCollection"
    Set dc1 = New DictCollection
    dc1.EmulateDictionary = True
    dc1.ZeroBasedIndex = False
    dc1.CompareMode = vbTextCompare
    dc1.ThrowErrors = False
    dc1.DefaultValueEnabled = False
    dc1.LazySorting = False
    Set dc2 = dc1.AddDC("key1")
    If dc2 Is Nothing Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc2.EmulateDictionary <> dc1.EmulateDictionary Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc2.ZeroBasedIndex <> dc1.ZeroBasedIndex Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc2.CompareMode <> dc1.CompareMode Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc2.ThrowErrors <> dc1.ThrowErrors Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc2.DefaultValueEnabled <> dc1.DefaultValueEnabled Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc2.LazySorting <> dc1.LazySorting Then logError DebugPrint, errors, ("-> " & currentTest)
    
currentTest = "[ADDDC-2] AddDC must be chainable"
    Set dc1 = New DictCollection
    On Error Resume Next
    dc1.AddDC(, 1).AddDC("key1").AddDC(, 0).AddDC ""
    If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.Item(0) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
    If Not IsObject(dc1.Item(1)) Then
        UtilAddArrayValue errors, ("-> " & currentTest)
    Else
        If TypeName(dc1.Item(1)) <> "DictCollection" Then
            UtilAddArrayValue errors, ("-> " & currentTest)
        Else
            ' 4 = filled key-value-store ("key1")
            If dc1.Item(1).CollectionType <> 4 Then logError DebugPrint, errors, ("-> " & currentTest)
            If Not IsObject(dc1.Item(1).Item("key1")) Then
                UtilAddArrayValue errors, ("-> " & currentTest)
            Else
                If TypeName(dc1.Item(1).Item("key1")) <> "DictCollection" Then
                    UtilAddArrayValue errors, ("-> " & currentTest)
                Else
                    ' 2 = filled array
                    If dc1.Item(1).Item("key1").CollectionType <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
                    If Not IsObject(dc1.Item(1).Item("key1").Item(0)) Then
                        UtilAddArrayValue errors, ("-> " & currentTest)
                    Else
                        If TypeName(dc1.Item(1).Item("key1").Item(0)) <> "DictCollection" Then
                            UtilAddArrayValue errors, ("-> " & currentTest)
                        Else
                            ' 2 = filled array
                            If dc1.Item(1).Item("key1").Item(0).CollectionType <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
                            If Not IsObject(dc1.Item(1).Item("key1").Item(0).Item(0)) Then
                                UtilAddArrayValue errors, ("-> " & currentTest)
                            Else
                                If TypeName(dc1.Item(1).Item("key1").Item(0).Item(0)) <> "DictCollection" Then
                                    UtilAddArrayValue errors, ("-> " & currentTest)
                                Else
                                    ' 1 = empty array
                                    If dc1.Item(1).Item("key1").Item(0).Item(0).CollectionType <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    On Error GoTo Fail: Err.Clear
currentTest = "[ADDDC-3] AddDC must throw correct errors"
    Set dc1 = New DictCollection: dc1.ThrowErrors = True: Set dc2 = New DictCollection: dc2.DefaultValueEnabled = False
    On Error Resume Next
    Err.Clear: dc1.AddDC UtilGetMissingValue: If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": must not throw error with key of type Missing")
    Err.Clear: dc1.AddDC 1: If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": must not throw error with key of type Integer")
    Err.Clear: dc1.AddDC 0.5: If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": must not throw error with key of type Double")
    Err.Clear: dc1.AddDC CDate(1): If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": must not throw error with key of type Date")
    testdata = Array(Array(1, 2), dc2, Null, CVErr(11), Nothing) ' invalid key types
    For i = 0 To UBound(testdata)
        Err.Clear: dc1.AddDC testdata(i): If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest & ": must throw error 13 with key of type " & TypeName(testdata(i)))
    Next
    On Error GoTo Fail: Err.Clear
currentTest = "[ADDDC-4] AddDC must work with .EmulateDictionary=True"
    Set dc1 = New DictCollection: dc1.EmulateDictionary = True: Set dc2 = New DictCollection: dc2.DefaultValueEnabled = False
    testdata = Array(CInt(1), CLng(2), CSng(3.5), CDbl(4.5), CCur(5), CDec(6), "", "x", UtilGetMissingValue, dc2, Null, CVErr(11), Nothing) ' valid Dictionary key types
    For i = 0 To UBound(testdata): dc1.AddDC(testdata(i)).Add "", i: Next
    For i = 0 To UBound(testdata)
        If Not dc1.Exists(testdata(i)) Then
            UtilAddArrayValue errors, ("-> " & currentTest)
        Else
            If TypeName(dc1.KeyOfItemAt(i)) <> TypeName(testdata(i)) Then logError DebugPrint, errors, ("-> " & currentTest)
            If Not IsObject(dc1.Item(testdata(i))) Then
                UtilAddArrayValue errors, ("-> " & currentTest)
            Else
                If dc1.Item(testdata(i)).ItemAt(0) <> i Then logError DebugPrint, errors, ("-> " & currentTest)
            End If
        End If
    Next
    On Error Resume Next
    ' try adding existing key
    Err.Clear: dc1.AddDC testdata(0): If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & ": adding an existing key must throw Error 457")
    On Error GoTo Fail: Err.Clear
    
currentTest = "[ADDDC-5] AddDC must work with .EmulateCollection=True"
    Set dc1 = New DictCollection: dc1.EmulateCollection = True: Set dc2 = New DictCollection: dc2.DefaultValueEnabled = False
    testdata = Array(Array(1, 2), dc2, Null, CVErr(11), Nothing) ' invalid key types
    On Error Resume Next
    For i = 0 To UBound(testdata)
        Err.Clear: dc1.AddDC testdata(i): If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest & ": adding keys of type " & TypeName(testdata(i)) & " must throw Error 13")
    Next
    ' try adding at index = 0 must not be possible
    Err.Clear: dc1.AddDC "b", 0: If Err.Number <> 9 Or dc1.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": adding at index 0 must throw Error 9")
    ' insert at first position
    Err.Clear: dc1.AddDC "", 1: If Err.Number <> 0 Or dc1.Count <> 1 Then logError DebugPrint, errors, ("-> " & currentTest & ": adding at first position must work")
    ' try adding duplicate key must not be possible
    Err.Clear: dc1.AddDC "": If Err.Number <> 457 Or dc1.Count <> 1 Then logError DebugPrint, errors, ("-> " & currentTest & ": adding duplicate key must throw Error 457")
    ' insert at last position
    Err.Clear: dc1.AddDC "a", -1: If Err.Number <> 0 Or dc1.Count <> 2 Then logError DebugPrint, errors, ("-> " & currentTest & ": adding at last position must work")
    Keys = dc1.Keys
    If Keys(0) <> "" Or Keys(1) <> "a" Then logError DebugPrint, errors, ("-> " & currentTest & ": must contain correct keys '' and 'a'")
    On Error GoTo Fail: Err.Clear
    
    GoTo ExitFunction
Fail:
    logError DebugPrint, errors, ("-> " & currentTest & " -> Error " & Err.Number & " " & Err.Description)
ExitFunction:
    testAdvancedFunctionality1 = errors
End Function

' Tests advanced functionality (DebugPrint=True -> print errors immediately)
Private Function testAdvancedFunctionality2(Optional DebugPrint As Boolean = False) As Variant
    On Error GoTo Fail: Err.Clear
    Dim errors As Variant, dc1 As DictCollection, dc2 As DictCollection, dc3 As DictCollection, currentTest As String
    Dim i As Long, j As Long, Items As Variant, Keys As Variant, testdata As Variant, val As Variant
    errors = Array()
    
currentTest = "[COPYTN-1] CopyItemsToNew must copy all settings if called without arguments"
    Set dc1 = New DictCollection:
    dc1.EmulateDictionary = True
    dc1.ThrowErrors = False
    dc1.DefaultValueEnabled = True
    dc1.ZeroBasedIndex = True
    dc1.EmptyCollectionValue = "a"
    dc1.NonExistingValue = "b"
    dc1.CompareMode = 2 ' database
    Set dc2 = dc1.CopyItemsToNew
    If dc2.ThrowErrors <> False Or dc2.DefaultValueEnabled <> True Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc2.ZeroBasedIndex <> True Then logError DebugPrint, errors, ("-> " & currentTest)
    If TypeName(dc2.EmptyCollectionValue) <> "String" Or dc2.EmptyCollectionValue <> "a" Then logError DebugPrint, errors, ("-> " & currentTest)
    If TypeName(dc2.NonExistingValue) <> "String" Or dc2.NonExistingValue <> "b" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc2.CompareMode <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[COPYTN-2] CopyItemsToNew must copy items and keys"
    Set dc1 = New DictCollection: dc1.Add , "A": dc1.Add "b", "B": dc1.Insert "C", 3, "c"
    Set dc2 = dc1.CopyItemsToNew
    Keys = dc2.Keys: If Keys(0) <> Empty Or Keys(1) <> "b" Or Keys(2) <> Empty Or Keys(3) <> "c" Then logError DebugPrint, errors, ("-> " & currentTest)
    Items = dc2.Items: If Items(0) <> "A" Or Items(1) <> "B" Or Items(2) <> Empty Or Items(3) <> "C" Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[COPYTN-3] CopyItemsToNew must copy object items"
    Set dc1 = New DictCollection: dc1.subDc(0).Add , 1: dc1.subDc(1).Add , 2:
    Set dc2 = dc1.CopyItemsToNew
    If dc2.AsDC(0).Item(0) <> 1 Or dc2.AsDC(1).Item(0) <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[COPYTN-4] CopyItemsToNew must work with .EmulateDictionary=True"
    Set dc1 = New DictCollection: dc1.EmulateDictionary = True
    testdata = Array(CInt(1), CLng(2), CSng(3.5), CDbl(4.5), CCur(5), CDec(6), "", "x", UtilGetMissingValue, dc2, Null, CVErr(11), Nothing) ' valid Dictionary key types
    For i = 0 To UBound(testdata): dc1.Add testdata(i), i: Next
    Set dc2 = dc1.CopyItemsToNew
    For i = 0 To UBound(testdata)
        If Not dc2.Exists(testdata(i)) Then
            UtilAddArrayValue errors, ("-> " & currentTest)
        Else
            If TypeName(dc2.KeyOfItemAt(i)) <> TypeName(testdata(i)) Then logError DebugPrint, errors, ("-> " & currentTest)
            If dc2.Item(testdata(i)) <> i Then UtilAddArrayValue errors, ("-> " & currentTest)
        End If
    Next

currentTest = "[COPYIT-1A] CopyItems must copy all settings if called with Nothing as argument"
    Set dc1 = New DictCollection: Set dc2 = Nothing
    dc1.EmulateDictionary = True
    dc1.ThrowErrors = True
    dc1.DefaultValueEnabled = True
    dc1.ZeroBasedIndex = False
    dc1.EmptyCollectionValue = "a"
    dc1.NonExistingValue = "b"
    dc1.CompareMode = 2 ' database
    dc1.CopyItems dc2
    If dc2.ThrowErrors <> True Or dc2.DefaultValueEnabled <> True Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc2.ZeroBasedIndex <> False Then logError DebugPrint, errors, ("-> " & currentTest)
    If TypeName(dc2.EmptyCollectionValue) <> "String" Or dc2.EmptyCollectionValue <> "a" Then logError DebugPrint, errors, ("-> " & currentTest)
    If TypeName(dc2.NonExistingValue) <> "String" Or dc2.NonExistingValue <> "b" Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc2.CompareMode <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[COPYIT-1B] CopyItems must not copy settings if called with existing DictCollection"
    Set dc2 = New DictCollection
    dc2.EmulateDictionary = False
    dc2.ThrowErrors = False
    dc2.DefaultValueEnabled = False
    dc2.ZeroBasedIndex = True
    dc2.EmptyCollectionValue = 1
    dc2.NonExistingValue = 2
    dc2.CompareMode = 1 ' text
    dc1.CopyItems dc2
    If dc2.ThrowErrors <> False Or dc2.DefaultValueEnabled <> False Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc2.ZeroBasedIndex <> True Then logError DebugPrint, errors, ("-> " & currentTest)
    If TypeName(dc2.EmptyCollectionValue) <> "Integer" Or dc2.EmptyCollectionValue <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
    If TypeName(dc2.NonExistingValue) <> "Integer" Or dc2.NonExistingValue <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc2.CompareMode <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[COPYIT-2] CopyItems must copy items and keys"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection: dc1.Add , "A": dc1.Add "b", "B": dc1.Insert "C", 3, "c"
    dc1.CopyItems dc2
    Keys = dc2.Keys: If Keys(0) <> Empty Or Keys(1) <> "b" Or Keys(2) <> Empty Or Keys(3) <> "c" Then logError DebugPrint, errors, ("-> " & currentTest)
    Items = dc2.Items: If Items(0) <> "A" Or Items(1) <> "B" Or Items(2) <> Empty Or Items(3) <> "C" Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[COPYIT-3] CopyItems must copy items and keys to given TargetIndex"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection: dc1.Add , "A": dc1.Add "b", "B": dc1.Insert "C", 3, "c"
    dc1.CopyItems dc2, 1
    Keys = dc2.Keys: If Keys(0) <> Empty Or Keys(1) <> Empty Or Keys(2) <> "b" Or Keys(3) <> Empty Or Keys(4) <> "c" Then logError DebugPrint, errors, ("-> " & currentTest)
    Items = dc2.Items: If Items(0) <> Empty Or Items(1) <> "A" Or Items(2) <> "B" Or Items(3) <> Empty Or Items(4) <> "C" Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[COPYIT-4] CopyItems must copy items and keys to given TargetIndex with TargetCollection.ZeroBasedIndex=False"
    Set dc1 = New DictCollection: dc1.Add , "A": dc1.Add "b", "B": dc1.Insert "C", 3, "c"
    Set dc2 = New DictCollection: dc2.ZeroBasedIndex = False
    dc1.CopyItems dc2, 2
    Keys = dc2.Keys: If Keys(0) <> Empty Or Keys(1) <> Empty Or Keys(2) <> "b" Or Keys(3) <> Empty Or Keys(4) <> "c" Then logError DebugPrint, errors, ("-> " & currentTest)
    Items = dc2.Items: If Items(0) <> Empty Or Items(1) <> "A" Or Items(2) <> "B" Or Items(3) <> Empty Or Items(4) <> "C" Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[COPYIT-5] CopyItems must copy object items"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection: dc1.subDc(0).Add , 1: dc1.subDc(1).Add , 2:
    dc1.CopyItems dc2
    If dc2.AsDC(0).Item(0) <> 1 Or dc2.AsDC(1).Item(0) <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    dc1.CopyItems dc2, 3 ' copy to index=3, leaves existing items untouched and inserts an empty item at index=2
    If dc2.AsDC(0).Item(0) <> 1 Or dc2.AsDC(1).Item(0) <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc2.Item(2) <> Empty Or dc2.AsDC(3).Item(0) <> 1 Or dc2.AsDC(4).Item(0) <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[COPYIT-6] CopyItems must work with .EmulateDictionary=True"
    Set dc1 = New DictCollection: Set dc2 = Nothing: dc1.EmulateDictionary = True
    testdata = Array(CInt(1), CLng(2), CSng(3.5), CDbl(4.5), CCur(5), CDec(6), "", "x", UtilGetMissingValue, New DictCollection, Null, CVErr(11), Nothing) ' valid Dictionary key types
    For i = 0 To UBound(testdata): dc1.Add testdata(i), i: Next
    dc1.CopyItems dc2
    For i = 0 To UBound(testdata)
        If Not dc2.Exists(testdata(i)) Then
            UtilAddArrayValue errors, ("-> " & currentTest)
        Else
            If TypeName(dc2.KeyOfItemAt(i)) <> TypeName(testdata(i)) Then logError DebugPrint, errors, ("-> " & currentTest)
            If dc2.Item(testdata(i)) <> i Then UtilAddArrayValue errors, ("-> " & currentTest)
        End If
    Next
currentTest = "[COPYIT-7] CopyItems must be chainable"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection: dc1.Add "a", 1
    Set dc2 = dc1.CopyItems(New DictCollection) ' get source DictCollection
    If Not IsObject(dc2) Then
        logError DebugPrint, errors, ("-> " & currentTest & ": function did not return object")
    Else
        If Not dc2 Is dc1 Then logError DebugPrint, errors, ("-> " & currentTest & ": returned object is not the same as source object")
    End If
    
    For i = 1 To 3 ' 1 = CloneToNew, 2 = CloneTo, 3 = CloneToPreserve
If i = 1 Then currentTest = "[CLONE-1A] CloneToNew must copy all settings correctly"
If i = 2 Then currentTest = "[CLONE-1B] CloneTo must copy all settings correctly"
If i = 3 Then currentTest = "[CLONE-1C] CloneToPreserve must not copy settings"
        Set dc1 = New DictCollection: Set dc2 = New DictCollection
        dc1.EmulateCollection = True
        dc1.DefaultValueEnabled = False
        dc1.EmptyCollectionValue = "empty"
        dc1.NonExistingValue = #1/1/2000#
        dc1.LazySorting = False
        dc1.ZeroBasedIndex = False
        dc1.CompareMode = vbTextCompare
        dc1.ThrowErrors = True
        ' clone
        If i = 1 Then Set dc2 = dc1.CloneToNew
        If i = 2 Then dc1.CloneTo dc2
        If i = 3 Then dc2.NonExistingValue = 1: dc2.LazySorting = True: dc1.CloneToPreserve dc2
        ' tests
        If i = 1 Or i = 2 Then
            If dc2.EmulateCollection <> True Or dc2.DefaultValueEnabled <> False Or dc2.EmptyCollectionValue <> "empty" Then logError DebugPrint, errors, ("-> " & currentTest)
            If dc2.NonExistingValue <> #1/1/2000# Or dc2.LazySorting <> False Or dc2.ZeroBasedIndex <> False Then logError DebugPrint, errors, ("-> " & currentTest)
            If dc2.CompareMode <> vbTextCompare Or dc2.ThrowErrors <> True Then logError DebugPrint, errors, ("-> " & currentTest)
        ElseIf i = 3 Then
            If dc2.NonExistingValue <> 1 Or dc2.LazySorting <> True Then logError DebugPrint, errors, ("-> " & currentTest)
        End If
If i = 1 Then currentTest = "[CLONE-2A] CloneToNew must copy all values"
If i = 2 Then currentTest = "[CLONE-2B] CloneTo must copy all values"
If i = 3 Then currentTest = "[CLONE-2C] CloneToPreserve must copy all values"
        Set dc1 = New DictCollection: Set dc2 = New DictCollection: Set dc3 = New DictCollection: dc3.DefaultValueEnabled = False
        testdata = Array(Array(Empty, Empty), Array("1", CInt(1)), Array("2", CLng(2)), Array("3", CSng(3.5)), Array("4", CDbl(4.5)), Array("5", CCur(5)), Array("6", CDec(6)), Array("7", "") _
        , Array("8", "x"), Array("9", dc3), Array("10", Null), Array("11", CVErr(11)), Array("12", Nothing)) ' valid DictColletion value types
        For j = 0 To UBound(testdata)
            dc1.Add testdata(j)(0), testdata(j)(1)
        Next
        ' clone
        If i = 1 Then Set dc2 = dc1.CloneToNew
        If i = 2 Then dc1.CloneTo dc2
        If i = 3 Then dc1.CloneToPreserve dc2
        ' test
        Items = dc2.Items: Keys = dc2.Keys
        For j = 0 To UBound(testdata)
            ' test items for equality
            Select Case TypeName(testdata(j)(1))
                Case "Nothing": If TypeName(Items(j)) <> "Nothing" Then logError DebugPrint, errors, ("-> " & currentTest)
                Case "DictCollection":
                    If TypeName(Items(j)) <> "DictCollection" Then
                        logError DebugPrint, errors, ("-> " & currentTest)
                    Else
                        ' cloned DictCollection has to be a different object
                        ' otherwise manipulating items in the original subcollection will affect the cloned one
                        If Items(j) Is testdata(j)(1) Then logError DebugPrint, errors, ("-> " & currentTest)
                    End If
                Case Else:
                    If TypeName(Items(j)) <> TypeName(testdata(j)(1)) Then
                        logError DebugPrint, errors, ("-> " & currentTest)
                    Else
                        If Items(j) <> testdata(j)(1) Then logError DebugPrint, errors, ("-> " & currentTest)
                    End If
            End Select
            ' test keys for equality
            If Keys(j) <> testdata(j)(0) Then logError DebugPrint, errors, ("-> " & currentTest)
        Next
If i = 1 Then currentTest = "[CLONE-3A] CloneToNew must replicate subcollection structure"
If i = 2 Then currentTest = "[CLONE-3B] CloneTo must replicate subcollection structure"
If i = 3 Then currentTest = "[CLONE-3C] CloneToPreserve must replicate subcollection structure"
        Set dc1 = New DictCollection: Set dc2 = New DictCollection: Set dc3 = Nothing
        dc1.AddDC("a").Add("a1", 1).Add "a2", 2 ' SubCollection "a" contains two values
        dc1.AddDC("b").Add("b1", 3).Add("b2", 4).AddDC "b3" ' SubCollection "b" contains two values and an empty SubCollection
        dc1.AddDC("c").AddDC.AddDC ' SubCollection "c" contains a SubCollection that contains an empty SubCollection
        ' clone
        If i = 1 Then Set dc2 = dc1.CloneToNew
        If i = 2 Then dc1.CloneTo dc2
        If i = 3 Then dc1.CloneToPreserve dc2
        ' test
        If Not dc2.Count = dc1.Count Then logError DebugPrint, errors, ("-> " & currentTest & ": cloned collection must have item count = " & dc1.Count)
        If Not dc2.KeyCount = dc1.KeyCount Then logError DebugPrint, errors, ("-> " & currentTest & ": cloned collection key count must be " & dc1.KeyCount)
        If TypeName(dc2(0)) <> "DictCollection" Or TypeName(dc2(1)) <> "DictCollection" Or TypeName(dc2(2)) <> "DictCollection" Then logError DebugPrint, errors, ("-> " & currentTest & ": cloned collection items must be of type DictCollection")
        If dc2.AsDC("a")("a1") <> 1 Or dc2.AsDC("a")("a2") <> 2 Then logError DebugPrint, errors, ("-> " & currentTest & ": cloned SubCollection 'a' items must have same values as source SubCollection")
        If dc2.AsDC("b")("b1") <> 3 Or dc2.AsDC("b")("b2") <> 4 Then logError DebugPrint, errors, ("-> " & currentTest & ": cloned SubCollection 'b' items must have same values as source SubCollection")
        If TypeName(dc2.AsDC("b")("b3")) <> "DictCollection" Then
            logError DebugPrint, errors, ("-> " & currentTest & ": cloned SubCollection 'b' item 'b3' must be of type SubCollection")
        Else
            If dc2.AsDC("b").AsDC("b3").Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": cloned SubCollection 'b' item 'b3' must be empty SubCollection")
        End If
        If TypeName(dc2.AsDC("c")(0)) <> "DictCollection" Then
            logError DebugPrint, errors, ("-> " & currentTest & ": cloned SubCollection 'c' must have SubCollection")
        Else
            If dc2.AsDC("c").AsDC(0).Count <> 1 Then
                logError DebugPrint, errors, ("-> " & currentTest & ": cloned SubCollection 'c' must have SubCollection with item count = 1")
            Else
                Set dc3 = dc2.AddDC("c").AsDC(0)(0)
                If TypeName(dc3) <> "DictCollection" Then
                    logError DebugPrint, errors, ("-> " & currentTest & ": cloned SubCollection 'c' first SubCollection must have SubCollection")
                Else
                    If dc3.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": cloned SubCollection 'c' first SubCollection must have SubCollection that is empty")
                End If
            End If
        End If
If i = 1 Then currentTest = "[CLONE-4A] CloneToNew must replicate all key types with .EmulateDictionary=True"
If i = 2 Then currentTest = "[CLONE-4B] CloneTo must replicate all key types with .EmulateDictionary=True"
If i = 3 Then currentTest = "[CLONE-4C] CloneToPreserve must replicate all key types with .EmulateDictionary=True"
        Set dc1 = New DictCollection: dc1.EmulateDictionary = True: Set dc2 = New DictCollection
        testdata = Array(CInt(1), CLng(2), CSng(3.5), CDbl(4.5), CCur(5), CDec(6), "", "x", UtilGetMissingValue, dc2, Null, CVErr(11), Nothing) ' valid Dictionary key types
        For j = 0 To UBound(testdata): dc1.Add testdata(j), j: Next
        ' clone
        If i = 1 Then Set dc2 = dc1.CloneToNew
        If i = 2 Then dc1.CloneTo dc2
        If i = 3 Then dc2.EmulateDictionary = True: dc1.CloneToPreserve dc2
        ' test
        If Not dc2.EmulateDictionary Then logError DebugPrint, errors, ("-> " & currentTest & ": cloned DictCollection must have .EmulateDictionary=True")
        For j = 0 To UBound(testdata)
            If Not dc2.Exists(testdata(j)) Then
                logError DebugPrint, errors, ("-> " & currentTest & ": cloned DictCollection must contain key '" & testdata(j) & "'")
            Else
                If TypeName(dc2.KeyOfItemAt(j)) <> TypeName(testdata(j)) Then logError DebugPrint, errors, ("-> " & currentTest & ": key at " & j & " must have type '" & TypeName(testdata(j)) & "'")
                If dc2.Item(testdata(j)) <> j Then logError DebugPrint, errors, ("-> " & currentTest & ": item at " & j & " must have value " & j)
            End If
        Next
    If i = 2 Then
currentTest = "[CLONE-5B] CloneTo must be chainable"
        Set dc1 = New DictCollection
        Set dc2 = dc1.CloneTo(New DictCollection)
        If dc2 Is Nothing Then
            logError DebugPrint, errors, ("-> " & currentTest)
        Else
            If Not dc2 Is dc1 Then logError DebugPrint, errors, ("-> " & currentTest)
        End If
    End If
    If i = 3 Then
currentTest = "[CLONE-5C] CloneToPreserve must be chainable"
        Set dc1 = New DictCollection
        Set dc2 = dc1.CloneToPreserve(New DictCollection)
        If dc2 Is Nothing Then
            logError DebugPrint, errors, ("-> " & currentTest)
        Else
            If Not dc2 Is dc1 Then logError DebugPrint, errors, ("-> " & currentTest)
        End If
currentTest = "[CLONE-6] CloneToPreserve must overwrite existing keys"
        Set dc1 = New DictCollection: dc1.Add "a", 1
        Set dc2 = New DictCollection: dc2.Add "a", 2
        dc1.CloneToPreserve dc2
        If dc2("a") <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[CLONE-7] CloneToPreserve must preserve existing keys"
        Set dc1 = New DictCollection: dc1.Add "a", 1
        Set dc2 = New DictCollection: dc2.Add "b", 2
        dc1.CloneToPreserve dc2
        If dc2.KeyOfItemAt(0) <> "b" Or dc2.ItemAt(0) <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc2.KeyOfItemAt(1) <> "a" Or dc2.ItemAt(1) <> 1 Then logError DebugPrint, errors, ("-> " & currentTest)
    End If
    
Next

    For i = 1 To 2 ' 1 = .Flatten(), 2 = .FlattenToNew()
If i = 1 Then currentTest = "[FLAT-1A] Flatten must replace empty and nonexisting DictCollections with .EmptyCollectionValue"
If i = 2 Then currentTest = "[FLAT-1B] FlattenToNew must replace empty and nonexisting DictCollections with .EmptyCollectionValue"
        Set dc1 = New DictCollection: Set dc2 = New DictCollection
        dc1.EmptyCollectionValue = "x"
        dc1.AddDC ' empty
        dc1.AddDC().CollectionType = 0 ' nonexisting
        If i = 1 Then
            dc1.Flatten: Items = dc1.Items: Keys = dc1.Keys
        Else
            Set dc2 = dc1.FlattenToNew: Items = dc2.Items: Keys = dc2.Keys ' test with TargetCollection
            If Not IsObject(dc1.Item(0)) Or Not IsObject(dc1.Item(1)) Then logError DebugPrint, errors, ("-> " & currentTest)
        End If
        If Items(0) <> dc1.EmptyCollectionValue Or Items(1) <> dc1.EmptyCollectionValue Then logError DebugPrint, errors, ("-> " & currentTest)
If i = 1 Then currentTest = "[FLAT-2A] Flatten must not change DictCollections with no SubCollections"
If i = 2 Then currentTest = "[FLAT-2B] FlattenToNew must not change DictCollections with no SubCollections"
        Set dc1 = New DictCollection: Set dc2 = New DictCollection
        dc1.Add , 1: dc1.Add "a", "A": dc1.Item(3) = "B"
        If i = 1 Then
            dc1.Flatten: Items = dc1.Items: Keys = dc1.Keys
        Else
            Set dc2 = dc1.FlattenToNew: Items = dc2.Items: Keys = dc2.Keys ' test with TargetCollection
            If dc1.KeyOfItemAt(0) <> Empty Or dc1.KeyOfItemAt(3) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
        End If
        If Keys(0) <> "_1" Or Keys(1) <> "a" Or Keys(2) <> "_3" Or Keys(3) <> "_4" Then logError DebugPrint, errors, ("-> " & currentTest)
        If Items(0) <> 1 Or Items(1) <> "A" Or Items(2) <> Empty Or Items(3) <> "B" Then logError DebugPrint, errors, ("-> " & currentTest)
If i = 1 Then currentTest = "[FLAT-3A] Flatten must insert items of all SubCollections"
If i = 2 Then currentTest = "[FLAT-3B] FlattenToNew must insert items of all SubCollections"
        Set dc1 = New DictCollection: Set dc2 = New DictCollection
        dc1.subDc("a").Add "key1", "item1"
        dc1.subDc("b").Add("key1", "item2").Add "key2", "item3"
        If i = 1 Then
            dc1.Flatten: Items = dc1.Items: Keys = dc1.Keys
        Else
            Set dc2 = dc1.FlattenToNew: Items = dc2.Items: Keys = dc2.Keys ' test with TargetCollection
            If dc1.Count <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
        End If
        If Keys(0) <> "a.key1" Or Keys(1) <> "b.key1" Or Keys(2) <> "b.key2" Then logError DebugPrint, errors, ("-> " & currentTest)
        If Items(0) <> "item1" Or Items(1) <> "item2" Or Items(2) <> "item3" Then logError DebugPrint, errors, ("-> " & currentTest)
If i = 1 Then currentTest = "[FLAT-4A] Flatten must change .EmulateDictionary to False"
If i = 2 Then currentTest = "[FLAT-4B] FlattenToNew must change .EmulateDictionary to False"
        Set dc1 = New DictCollection: dc1.EmulateDictionary = True: Set dc2 = New DictCollection: dc2.DefaultValueEnabled = False
        dc1.Add CInt(1), 1: dc1.Add UtilGetMissingValue, "A": dc1.Add CVErr(11), Empty: dc1.Add dc2, "B"
        If i = 1 Then
            dc1.Flatten: Items = dc1.Items: Keys = dc1.Keys
            If dc1.EmulateDictionary Then logError DebugPrint, errors, ("-> " & currentTest)
        Else
            Set dc2 = dc1.FlattenToNew: Items = dc2.Items: Keys = dc2.Keys ' test with TargetCollection
            If Not dc1.EmulateDictionary Then logError DebugPrint, errors, ("-> " & currentTest)
            If Not IsError(dc1.KeyOfItemAt(2)) Or Not IsObject(dc1.KeyOfItemAt(3)) Then logError DebugPrint, errors, ("-> " & currentTest)
        End If
        If TypeName(Keys(0)) = "Integer" Or IsMissing(Keys(1)) Or IsError(Keys(2)) Or IsObject(Keys(3)) Then logError DebugPrint, errors, ("-> " & currentTest)
        If Items(0) <> 1 Or Items(1) <> "A" Or Items(2) <> Empty Or Items(3) <> "B" Then logError DebugPrint, errors, ("-> " & currentTest)
If i = 1 Then currentTest = "[FLAT-5A] Flatten must assign correct keys to all items (logic of .EnsureAllItemsHaveKeys)"
If i = 2 Then currentTest = "[FLAT-5B] FlattenToNew must assign correct keys to all items (logic of .EnsureAllItemsHaveKeys)"
        Set dc1 = New DictCollection: Set dc2 = New DictCollection
        dc1.AddDC().AddDC().Add(, "item1").Add "_1", "item2"
        dc1.AddDC("_1").Add , "item3"
        dc1.AddDC().Add , "item4"
        dc1.AddDC("_3").Add , "item5"
        dc1.AddDC("_3_1").Add(, "item6").Add , "item7"
        If i = 1 Then
            dc1.Flatten: Items = dc1.Items: Keys = dc1.Keys
        Else
            Set dc2 = dc1.FlattenToNew: Items = dc2.Items: Keys = dc2.Keys ' test with TargetCollection
            If dc1.Count <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
        End If
        If Items(0) <> "item1" Or Keys(0) <> "_1_1._1._1_1" Then logError DebugPrint, errors, ("-> " & currentTest)
        If Items(1) <> "item2" Or Keys(1) <> "_1_1._1._1" Then logError DebugPrint, errors, ("-> " & currentTest)
        If Items(2) <> "item3" Or Keys(2) <> "_1._1" Then logError DebugPrint, errors, ("-> " & currentTest)
        If Items(3) <> "item4" Or Keys(3) <> "_3_2._1" Then logError DebugPrint, errors, ("-> " & currentTest)
        If Items(4) <> "item5" Or Keys(4) <> "_3._1" Then logError DebugPrint, errors, ("-> " & currentTest)
        If Items(5) <> "item6" Or Keys(5) <> "_3_1._1" Then logError DebugPrint, errors, ("-> " & currentTest)
        If Items(6) <> "item7" Or Keys(6) <> "_3_1._2" Then logError DebugPrint, errors, ("-> " & currentTest)
If i = 1 Then
    currentTest = "[FLAT-6] Flatten must be chainable"
        Set dc1 = New DictCollection: Set dc2 = New DictCollection
        Set dc2 = dc1.Flatten
        If dc2 Is Nothing Then
            logError DebugPrint, errors, ("-> " & currentTest)
        Else
            If Not dc2 Is dc1 Then logError DebugPrint, errors, ("-> " & currentTest)
        End If
End If
    Next

    For i = 1 To 2 ' 1 = .Unflatten(), 2 = .UnflattenToNew()
If i = 1 Then currentTest = "[UNFL-1A] Unflatten must restore array items in correct order"
If i = 2 Then currentTest = "[UNFL-1B] UnflattenToNew must restore array items in correct order"
        Set dc1 = New DictCollection: dc1.Add "_3", "item3": dc1.Add "_1", "item1": dc1.Add "_2", "item2"
        If i = 1 Then dc1.Unflatten Else Set dc1 = dc1.UnflattenToNew
        If dc1.CollectionType <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
        Items = dc1.Items: If Items(0) <> "item1" Or Items(1) <> "item2" Or Items(2) <> "item3" Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc1.KeyCount <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
If i = 1 Then currentTest = "[UNFL-2A] Unflatten must restore array items in correct order and leave non-unflattable keys intact"
If i = 2 Then currentTest = "[UNFL-2B] UnflattenToNew must restore array items in correct order and leave non-unflattable keys intact"
        Set dc1 = New DictCollection: dc1.Add "_3", "item3": dc1.Add "_1_x", "item4": dc1.Add "_1", "item1": dc1.Add "_2", "item2": dc1.Add "_1_", "item5"
        If i = 1 Then dc1.Unflatten Else Set dc1 = dc1.UnflattenToNew
        If dc1.Count <> 5 Or dc1.KeyCount <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
        Items = dc1.Items: If Items(0) <> "item1" Or Items(1) <> "item2" Or Items(2) <> "item3" Or Items(3) <> "item4" Or Items(4) <> "item5" Then logError DebugPrint, errors, ("-> " & currentTest)
        Keys = dc1.Keys(False): If Keys(0) <> "_1_x" Or Keys(1) <> "_1_" Then logError DebugPrint, errors, ("-> " & currentTest)
If i = 1 Then currentTest = "[UNFL-3A] Unflatten must restore subcollection and its items"
If i = 2 Then currentTest = "[UNFL-3B] UnflattenToNew must restore subcollection and its items"
        Set dc1 = New DictCollection: dc1.Add "a.b", "B": dc1.Add "b.e", "E": dc1.Add "a.a", "A": dc1.Add "b._1", "D": dc1.Add "a.c", "C"
        If i = 1 Then dc1.Unflatten Else Set dc1 = dc1.UnflattenToNew
        ' desired result -> { a{b{B},a{A},c{C}} , b{D,e{E} }
        If dc1.Count <> 2 Or dc1.KeyCount <> 2 Then logError DebugPrint, errors, ("-> " & currentTest)
        Items = dc1.AsDC(0).Items: If Items(0) <> "B" Or Items(1) <> "A" Or Items(2) <> "C" Then logError DebugPrint, errors, ("-> " & currentTest)
        Keys = dc1.AsDC(0).Keys: If Keys(0) <> "b" Or Keys(1) <> "a" Or Keys(2) <> "c" Then logError DebugPrint, errors, ("-> " & currentTest)
        Items = dc1.AsDC(1).Items: If Items(0) <> "D" Or Items(1) <> "E" Then logError DebugPrint, errors, ("-> " & currentTest)
        Keys = dc1.AsDC(1).Keys: If Keys(0) <> Empty Or Keys(1) <> "e" Then logError DebugPrint, errors, ("-> " & currentTest)
If i = 1 Then currentTest = "[UNFL-4A] Unflatten must create correct subtree structure"
If i = 2 Then currentTest = "[UNFL-4B] UnflattenToNew must create correct subtree structure"
        ' this is the reverse of [FLAT-5]
        Set dc1 = New DictCollection
        dc1.Add "_1_1._1._1_1", "item1"
        dc1.Add "_1_1._1._1", "item2"
        dc1.Add "_1._1", "item3"
        dc1.Add "_3_2._1", "item4"
        dc1.Add "_3._1", "item5"
        dc1.Add "_3_1._1", "item6"
        dc1.Add "_3_1._2", "item7"
        If i = 1 Then dc1.Unflatten Else Set dc1 = dc1.UnflattenToNew
        ' desired result -> { {{item1,item2}} , _1{item3} , {item4} , _3{item5} , _3_1{item6,item7} }
        If dc1.Count <> 5 Or dc1.KeyCount <> 3 Then logError DebugPrint, errors, ("-> " & currentTest)
        Items = dc1.AsDC(0).AsDC(0).Items: If Items(0) <> "item1" Or Items(1) <> "item2" Then logError DebugPrint, errors, ("-> " & currentTest)
        Items = dc1.AsDC("_1").Items: If Items(0) <> "item3" Then logError DebugPrint, errors, ("-> " & currentTest)
        Items = dc1.AsDC(2).Items: If Items(0) <> "item4" Then logError DebugPrint, errors, ("-> " & currentTest)
        Items = dc1.AsDC("_3").Items: If Items(0) <> "item5" Then logError DebugPrint, errors, ("-> " & currentTest)
        Items = dc1.AsDC("_3_1").Items: If Items(0) <> "item6" Or Items(1) <> "item7" Then logError DebugPrint, errors, ("-> " & currentTest)
If i = 1 Then
    currentTest = "[UNFL-5] Unflatten must be chainable"
        Set dc1 = New DictCollection: Set dc2 = New DictCollection
        Set dc2 = dc1.Unflatten
        If dc2 Is Nothing Then
            logError DebugPrint, errors, ("-> " & currentTest)
        Else
            If Not dc2 Is dc1 Then logError DebugPrint, errors, ("-> " & currentTest)
        End If
End If
    Next
    
    For i = 1 To 4
        Set dc1 = New DictCollection: Set dc2 = New DictCollection
If i = 1 Then currentTest = "[MATCHKNEW-1A] MatchKeysToNew must copy settings if nothing to drop and empty source DictCollection"
If i = 2 Then dc1.Add "x", 1: currentTest = "[MATCHKNEW-1B] MatchKeysToNew must copy settings if nothing to drop and filled source DictCollection"
If i = 3 Then Keys = Array("x"): currentTest = "[MATCHKNEW-1C] MatchKeysToNew must copy settings if empty source DictCollection"
If i = 4 Then dc1.Add "x", 1: Keys = Array("x"): currentTest = "[MATCHKNEW-1D] MatchKeysToNew must copy settings if filled source DictCollection"
        dc1.EmulateCollection = True
        dc1.DefaultValueEnabled = False
        dc1.EmptyCollectionValue = "empty"
        dc1.NonExistingValue = #1/1/2000#
        dc1.LazySorting = False
        dc1.ZeroBasedIndex = False
        dc1.CompareMode = vbTextCompare
        dc1.ThrowErrors = True
        ' filter
        Set dc2 = dc1.MatchKeysToNew(Array())
        ' tests
        If dc2.EmulateCollection <> True Or dc2.DefaultValueEnabled <> False Or dc2.EmptyCollectionValue <> "empty" Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc2.NonExistingValue <> #1/1/2000# Or dc2.LazySorting <> False Or dc2.ZeroBasedIndex <> False Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc2.CompareMode <> vbTextCompare Or dc2.ThrowErrors <> True Then logError DebugPrint, errors, ("-> " & currentTest)
    Next
currentTest = "[MATCHKNEW-2] MatchKeysToNew must return empty DictCollection if there is nothing to match"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.Add("a", 1).Add("b", 2).Add(, 2.5).Add "c", 3
    Set dc2 = dc1.MatchKeysToNew(Array())
    If dc2.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection must be empty")
currentTest = "[MATCHKNEW-3A] MatchKeysToNew must find exact matches with case-sensitive matching"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.Add("a", 1).Add("b", 2).Add(, 2.5).Add "c", 3
    Set dc2 = dc1.MatchKeysToNew(Array("a", "B"), "?", vbBinaryCompare)
    If dc2.Count <> 2 Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection has icorrect item count")
    If dc2.KeyOfItemAt(0) <> "a" Or dc2.ItemOf(0) = 1 Or dc2.KeyOfItemAt(1) <> "B" Or dc2.ItemOf(1) = "?" Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection contains incorrect keys/items")
currentTest = "[MATCHKNEW-3B] MatchKeysToNew must find matches with case-insensitive matching"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.Add("a", 1).Add("b", 2).Add(, 2.5).Add "c", 3
    Set dc2 = dc1.MatchKeysToNew(Array("a", "B"), , vbTextCompare)
    If dc2.Count <> 2 Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection has incorrect item count")
    If dc2.KeyOfItemAt(0) <> "a" Or dc2.ItemOf(0) = 1 Or dc2.KeyOfItemAt(1) <> "B" Or dc2.ItemOf(1) = 2 Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection contains incorrect keys/items")
currentTest = "[MATCHKNEW-4] MatchKeysToNew must return all keys if nothing was found"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.Add("x", 1).Add("_X_", 2).Add("_x", 3).Add "X_", 4
    Set dc2 = dc1.MatchKeysToNew(Array("a", "b", "c"), -1, vbBinaryCompare)
    If dc2.Count <> 3 Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection has incorrect item count")
    If dc2.KeyOfItemAt(0) <> "a" Or dc2.ItemOf(0) = -1 Or dc2.KeyOfItemAt(1) <> "b" Or dc2.ItemOf(1) = -1 Or dc2.KeyOfItemAt(2) <> "c" Or dc2.ItemOf(2) = -1 Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection contains incorrect keys/items")
currentTest = "[MATCHKNEW-5] MatchKeysToNew must clone matching SubCollections"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.AddDC("x").Add "a", 1
    dc1.AddDC("y").Add "b", 2
    Set dc2 = dc1.MatchKeysToNew(Array("y")) ' match SubCollection 'y'
    dc1.AsDC("y").RemoveAll ' empty source SubCollection, must not affect cloned SubCollection
    If dc2.Count <> 1 Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection has incorrect item number")
    If dc2.AsDC("y").Count <> 1 Or dc2.AsDC("y").KeyOfItemAt(0) <> "b" Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection has incorrect SubCollection")
currentTest = "[MATCHKNEW-6] MatchKeysToNew must trow correct error if KeyArr is of invalid type"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.ThrowErrors = True
    testdata = Array(CInt(1), CDbl(1#), "text", New DictCollection)
    On Error Resume Next
    For i = 0 To UBound(testdata)
        Err.Clear: Set dc2 = dc1.MatchKeysToNew(testdata(i)): If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest & ": KeyArr of Type " & TypeName(testdata(i)) & " must throw Error 13")
    Next
    On Error GoTo Fail: Err.Clear
       
    For i = 1 To 4
        Set dc1 = New DictCollection: Set dc2 = New DictCollection
If i = 1 Then currentTest = "[KEEPKNEW-1A] KeepKeysToNew must copy settings if nothing to drop and empty source DictCollection"
If i = 2 Then dc1.Add "x", 1: currentTest = "[KEEPKNEW-1B] KeepKeysToNew must copy settings if nothing to drop and filled source DictCollection"
If i = 3 Then Keys = Array("x"): currentTest = "[KEEPKNEW-1C] KeepKeysToNew must copy settings if empty source DictCollection"
If i = 4 Then dc1.Add "x", 1: Keys = Array("x"): currentTest = "[KEEPKNEW-1D] KeepKeysToNew must copy settings if filled source DictCollection"
        dc1.EmulateCollection = True
        dc1.DefaultValueEnabled = False
        dc1.EmptyCollectionValue = "empty"
        dc1.NonExistingValue = #1/1/2000#
        dc1.LazySorting = False
        dc1.ZeroBasedIndex = False
        dc1.CompareMode = vbTextCompare
        dc1.ThrowErrors = True
        ' filter
        Set dc2 = dc1.KeepKeysToNew(Array())
        ' tests
        If dc2.EmulateCollection <> True Or dc2.DefaultValueEnabled <> False Or dc2.EmptyCollectionValue <> "empty" Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc2.NonExistingValue <> #1/1/2000# Or dc2.LazySorting <> False Or dc2.ZeroBasedIndex <> False Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc2.CompareMode <> vbTextCompare Or dc2.ThrowErrors <> True Then logError DebugPrint, errors, ("-> " & currentTest)
    Next
currentTest = "[KEEPKNEW-2] KeepKeysToNew must return empty DictCollection if there is nothing to keep"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.Add("a", 1).Add("b", 2).Add(, 2.5).Add "c", 3
    Set dc2 = dc1.KeepKeysToNew(Array())
    If dc2.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection must be empty")
currentTest = "[KEEPKNEW-3A] KeepKeysToNew must keep exact matches with case-sensitive matching"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.Add("a", 1).Add("b", 2).Add(, 2.5).Add "c", 3
    Set dc2 = dc1.KeepKeysToNew(Array("a", "B"), , vbBinaryCompare)
    If dc2.Count <> 1 Or dc2.KeyOfItemAt(0) <> "a" Or dc2.ItemOf(0) = 1 Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection contains incorrect key/item")
currentTest = "[KEEPKNEW-3B] KeepKeysToNew must keep exact matches with case-insensitive matching"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.Add("a", 1).Add("b", 2).Add(, 2.5).Add "c", 3
    Set dc2 = dc1.KeepKeysToNew(Array("a", "B"), , vbTextCompare)
    If dc2.Count <> 2 Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection has incorrect item count")
    If dc2.KeyAtKeyIndex(0) <> "a" Or dc2.KeyAtKeyIndex(1) <> "b" Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection keys are incorrect")
    If dc2.ItemAt(0) <> 1 Or dc2.ItemAt(1) <> 2 Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection items are incorrect")
currentTest = "[KEEPKNEW-4A] KeepKeysToNew must keep matching wildcards with case-sensitive matching"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.Add("x", 1).Add("_X_", 2).Add("_x", 3).Add "X_", 4
    ' keep everything that contains "X"
    Set dc2 = dc1.KeepKeysToNew(Array("*X*"), True, vbBinaryCompare)
    If dc2.Count <> 2 Or dc2.KeyOfItemAt(0) <> "_X_" Or dc2.KeyOfItemAt(1) <> "X_" Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection has incorrect keys")
    ' keep everything that starts with "X"
    Set dc2 = dc1.KeepKeysToNew(Array("X*"), True, vbBinaryCompare)
    If dc2.Count <> 1 Or dc2.KeyOfItemAt(0) <> "X_" Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection has incorrect keys")
    ' keep everything that ends with "x"
    Set dc2 = dc1.KeepKeysToNew(Array("*x"), True, vbBinaryCompare)
    If dc2.Count <> 2 Or dc2.KeyOfItemAt(0) <> "x" Or dc2.KeyOfItemAt(1) <> "_x" Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection has incorrect keys")
currentTest = "[KEEPKNEW-4B] KeepKeysToNew must keep matching wildcards with case-insensitive matching"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.Add("x", 1).Add("_X_", 2).Add("_X", 3).Add "X_", 4
    ' keep everything that contains "X" or "x"
    Set dc2 = dc1.KeepKeysToNew(Array("*X*"), True, vbTextCompare)
    If dc2.Count <> 4 Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection has incorrect keys")
    ' keep everything that starts with "X" or "x"
    Set dc2 = dc1.KeepKeysToNew(Array("x*"), True, vbTextCompare)
    If dc2.Count <> 2 Or dc2.KeyOfItemAt(0) <> "x" Or dc2.KeyOfItemAt(1) <> "X_" Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection has incorrect keys")
    ' keep everything that ends with "X" or "x" or matches "_*_"
    Set dc2 = dc1.KeepKeysToNew(Array("*x", "_*_"), True, vbTextCompare)
    If dc2.Count <> 3 Or dc2.KeyOfItemAt(0) <> "x" Or dc2.KeyOfItemAt(1) <> "_X_" Or dc2.KeyOfItemAt(2) <> "_X" Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection has incorrect keys")
currentTest = "[KEEPKNEW-5] KeepKeysToNew must clone non-matching SubCollections"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.AddDC("x").Add "a", 1
    dc1.AddDC("y").Add "b", 2
    Set dc2 = dc1.KeepKeysToNew(Array("y")) ' keep SubCollection 'y'
    dc1.AsDC("y").RemoveAll ' empty source SubCollection, must not affect cloned SubCollection
    If dc2.Count <> 1 Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection has incorrect item number")
    If dc2.AsDC("y").Count <> 1 Or dc2.AsDC("y").KeyOfItemAt(0) <> "b" Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection has incorrect SubCollection")
currentTest = "[KEEPKNEW-6] KeepKeysToNew must trow correct error if KeyArr is of invalid type"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.ThrowErrors = True
    testdata = Array(CInt(1), CDbl(1#), "text", New DictCollection)
    On Error Resume Next
    For i = 0 To UBound(testdata)
        Err.Clear: Set dc2 = dc1.KeepKeysToNew(testdata(i)): If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest & ": KeyArr of Type " & TypeName(testdata(i)) & " must throw Error 13")
    Next
    On Error GoTo Fail: Err.Clear
    
    For i = 1 To 4
        Set dc1 = New DictCollection: Set dc2 = New DictCollection
If i = 1 Then currentTest = "[DROPKNEW-1A] DropKeysToNew must copy settings if nothing to drop and empty source DictCollection"
If i = 2 Then dc1.Add "x", 1: currentTest = "[DROPKNEW-1B] DropKeysToNew must copy settings if nothing to drop and filled source DictCollection"
If i = 3 Then Keys = Array("x"): currentTest = "[DROPKNEW-1C] DropKeysToNew must copy settings if empty source DictCollection"
If i = 4 Then dc1.Add "x", 1: Keys = Array("x"): currentTest = "[DROPKNEW-1D] DropKeysToNew must copy settings if filled source DictCollection"
        dc1.EmulateCollection = True
        dc1.DefaultValueEnabled = False
        dc1.EmptyCollectionValue = "empty"
        dc1.NonExistingValue = #1/1/2000#
        dc1.LazySorting = False
        dc1.ZeroBasedIndex = False
        dc1.CompareMode = vbTextCompare
        dc1.ThrowErrors = True
        ' filter
        Set dc2 = dc1.DropKeysToNew(Array())
        ' tests
        If dc2.EmulateCollection <> True Or dc2.DefaultValueEnabled <> False Or dc2.EmptyCollectionValue <> "empty" Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc2.NonExistingValue <> #1/1/2000# Or dc2.LazySorting <> False Or dc2.ZeroBasedIndex <> False Then logError DebugPrint, errors, ("-> " & currentTest)
        If dc2.CompareMode <> vbTextCompare Or dc2.ThrowErrors <> True Then logError DebugPrint, errors, ("-> " & currentTest)
    Next
currentTest = "[DROPKNEW-2] DropKeysToNew must clone all items with keys if nothing to drop"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.Add("a", 1).Add("b", 2).Add(, 2.5).Add "c", 3
    ' filter
    Set dc2 = dc1.DropKeysToNew(Array())
    ' tests
    If dc2.Count <> 3 Or dc2.KeyCount <> 3 Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection .Count=" & dc2.Count & " must be " & 3 & " and .KeyCount=" & dc2.KeyCount & " must be 3")
    If dc2.KeyAtKeyIndex(0) <> "a" Or dc2.KeyAtKeyIndex(1) <> "b" Or dc2.KeyAtKeyIndex(2) <> "c" Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection keys are incorrect")
    If dc2.ItemAt(0) <> 1 Or dc2.ItemAt(1) <> 2 Or dc2.ItemAt(2) <> 3 Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection items are incorrect")
currentTest = "[DROPKNEW-3A] DropKeysToNew must drop exact matches with case-sensitive matching"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.Add("a", 1).Add("b", 2).Add(, 2.5).Add "c", 3
    Set dc2 = dc1.DropKeysToNew(Array("B"), , vbBinaryCompare)
    If dc2.Count <> 3 Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection .Count=" & dc2.Count & " must be " & 3)
    If dc2.KeyAtKeyIndex(0) <> "a" Or dc2.KeyAtKeyIndex(1) <> "b" Or dc2.KeyAtKeyIndex(2) <> "c" Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection keys are incorrect")
    If dc2.ItemAt(0) <> 1 Or dc2.ItemAt(1) <> 2 Or dc2.ItemAt(2) <> 3 Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection items are incorrect")
currentTest = "[DROPKNEW-3B] DropKeysToNew must drop exact matches with case-insensitive matching"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.Add("a", 1).Add("b", 2).Add(, 2.5).Add "c", 3
    Set dc2 = dc1.DropKeysToNew(Array("B"), , vbTextCompare)
    If dc2.Count <> 2 Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection .Count=" & dc2.Count & " must be " & 2)
    If dc2.KeyAtKeyIndex(0) <> "a" Or dc2.KeyAtKeyIndex(1) <> "c" Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection keys are incorrect")
    If dc2.ItemAt(0) <> 1 Or dc2.ItemAt(1) <> 3 Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection items are incorrect")
currentTest = "[DROPKNEW-4A] DropKeysToNew must drop matching wildcards with case-sensitive matching"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.Add("x", 1).Add("_X_", 2).Add("_x", 3).Add "X_", 4
    ' drop everything that contains "X"
    Set dc2 = dc1.DropKeysToNew(Array("*X*"), True, vbBinaryCompare)
    If dc2.Count <> 2 Or dc2.KeyOfItemAt(0) <> "x" Or dc2.KeyOfItemAt(1) <> "_x" Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection has incorrect keys")
    ' drop everything that starts with "X"
    Set dc2 = dc1.DropKeysToNew(Array("X*"), True, vbBinaryCompare)
    If dc2.Count <> 3 Or dc2.KeyOfItemAt(0) <> "x" Or dc2.KeyOfItemAt(1) <> "_X_" Or dc2.KeyOfItemAt(2) <> "_x" Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection has incorrect keys")
    ' drop everything that ends with "x"
    Set dc2 = dc1.DropKeysToNew(Array("*x"), True, vbBinaryCompare)
    If dc2.Count <> 2 Or dc2.KeyOfItemAt(0) <> "_X_" Or dc2.KeyOfItemAt(1) <> "X_" Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection has incorrect keys")
currentTest = "[DROPKNEW-4B] DropKeysToNew must drop matching wildcards with case-insensitive matching"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.Add("x", 1).Add("_X_", 2).Add("_X", 3).Add "X_", 4
    ' drop everything that contains "X" or "x"
    Set dc2 = dc1.DropKeysToNew(Array("*X*"), True, vbTextCompare)
    If dc2.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection must be empty")
    ' drop everything that starts with "X" or "x"
    Set dc2 = dc1.DropKeysToNew(Array("x*"), True, vbTextCompare)
    If dc2.Count <> 2 Or dc2.KeyOfItemAt(0) <> "_X_" Or dc2.KeyOfItemAt(1) <> "_X" Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection has incorrect keys")
    ' drop everything that ends with "X" or "x" or matches "_*_"
    Set dc2 = dc1.DropKeysToNew(Array("*x", "_*_"), True, vbTextCompare)
    If dc2.Count <> 1 Or dc2.KeyOfItemAt(0) <> "X_" Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection has incorrect keys")
currentTest = "[DROPKNEW-5] DropKeysToNew must clone non-matching SubCollections"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.AddDC("x").Add "a", 1
    dc1.AddDC("y").Add "b", 2
    Set dc2 = dc1.DropKeysToNew(Array("x")) ' drop SubCollection 'x'
    dc1.AsDC("y").RemoveAll ' empty source SubCollection, must not affect cloned SubCollection
    If dc2.Count <> 1 Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection has incorrect item number")
    If dc2.AsDC("y").Count <> 1 Or dc2.AsDC("y").KeyOfItemAt(0) <> "b" Then logError DebugPrint, errors, ("-> " & currentTest & ": new DictCollection has incorrect SubCollection")
currentTest = "[DROPKNEW-6] DropKeysToNew must trow correct error if KeyArr is of invalid type"
    Set dc1 = New DictCollection: Set dc2 = New DictCollection
    dc1.ThrowErrors = True
    testdata = Array(CInt(1), CDbl(1#), "text", New DictCollection)
    On Error Resume Next
    For i = 0 To UBound(testdata)
        Err.Clear: Set dc2 = dc1.DropKeysToNew(testdata(i)): If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest & ": KeyArr of Type " & TypeName(testdata(i)) & " must throw Error 13")
    Next
    On Error GoTo Fail: Err.Clear
    
    GoTo ExitFunction
Fail:
    logError DebugPrint, errors, ("-> " & currentTest & " -> Error " & Err.Number & " " & Err.Description)
ExitFunction:
    testAdvancedFunctionality2 = errors
End Function


' tests utilities (DebugPrint=True -> print errors immediately)
Private Function testUtilFunctions(Optional DebugPrint As Boolean = False) As Variant
    On Error GoTo Fail: Err.Clear
    Dim errors As Variant, currentTest As String, i As Long, j As Long, testdata As Variant, val As Variant
    errors = Array()

currentTest = "[USMW-1] UtilStringMatchesWildcard must not match with empty string"
    If UtilStringMatchesWildcard("", Split("*x*", "*")) Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[USMW-2] UtilStringMatchesWildcard must not match with empty wildcard"
    If UtilStringMatchesWildcard("a", Split("", "*")) Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[USMW-3] UtilStringMatchesWildcard must not match when searchtext is nonexistent"
    If UtilStringMatchesWildcard("abc", Split("x", "*")) Then logError DebugPrint, errors, ("-> " & currentTest)
    If UtilStringMatchesWildcard("abc", Split("*x", "*")) Then logError DebugPrint, errors, ("-> " & currentTest)
    If UtilStringMatchesWildcard("abc", Split("x*", "*")) Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[USMW-4] UtilStringMatchesWildcard must not match when searchtext is not at start"
    If UtilStringMatchesWildcard("abc", Split("b*", "*")) Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[USMW-5] UtilStringMatchesWildcard must not match when searchtext is not at end"
    If UtilStringMatchesWildcard("abc", Split("*b", "*")) Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[USMW-6] UtilStringMatchesWildcard must not match when searchtext is not exactly matched"
    If UtilStringMatchesWildcard("bc", Split("b", "*")) Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[USMW-7] UtilStringMatchesWildcard must match when searchtext is at start"
    If Not UtilStringMatchesWildcard("abc", Split("a*", "*")) Then logError DebugPrint, errors, ("-> " & currentTest)
    If Not UtilStringMatchesWildcard("abc", Split("*a*", "*")) Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[USMW-8] UtilStringMatchesWildcard must match when searchtext is at end"
    If Not UtilStringMatchesWildcard("abc", Split("*c", "*")) Then logError DebugPrint, errors, ("-> " & currentTest)
    If Not UtilStringMatchesWildcard("abc", Split("*c*", "*")) Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[USMW-9] UtilStringMatchesWildcard must match when searchtext is in between"
    If Not UtilStringMatchesWildcard("abc", Split("*b*", "*")) Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[USMW-10] UtilStringMatchesWildcard must match when searchtext matches exactly"
    If Not UtilStringMatchesWildcard("abc", Split("abc", "*")) Then logError DebugPrint, errors, ("-> " & currentTest)
    

    GoTo ExitFunction
Fail:
    logError DebugPrint, errors, ("-> " & currentTest & " -> Error " & Err.Number & " " & Err.Description)
ExitFunction:
    testUtilFunctions = errors
End Function


' tests bugfixes, workarounds and special cases (DebugPrint=True -> print errors immediately)
Private Function testFixes(Optional DebugPrint As Boolean = False) As Variant
    On Error GoTo Fail: Err.Clear
    Dim errors As Variant, dc1 As DictCollection, dc2 As DictCollection, currentTest As String, i As Long, j As Long, Items As Variant, Keys As Variant, testdata As Variant, val As Variant
    Dim intItems() As Variant, intKeys() As String
    errors = Array()

currentTest = "[EDPL-1] Endless Default Property Evaluation Loop must not be possible": Set dc1 = New DictCollection: dc1.DefaultValueEnabled = False
    If dc1 <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    If dc1.Item(UtilGetMissingValue) <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error Resume Next
    Err.Clear: Set dc2 = dc1: If Not dc2 Is dc1 Or Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: testdata = Array(dc1): If Err.Number <> 0 Or testdata(0) <> dc1.NonExistingValue Then logError DebugPrint, errors, ("-> " & currentTest)
    On Error GoTo Fail: Err.Clear

    GoTo ExitFunction
Fail:
    logError DebugPrint, errors, ("-> " & currentTest & " -> Error " & Err.Number & " " & Err.Description)
ExitFunction:
    testFixes = errors
End Function



' ======================== END: FUNCTIONALITY TESTS =================================================================================================




' ======================== START: COMPATIBILITY TESTS ===============================================================================================
Public Function TestCompatibility(Optional DebugPrint As Boolean = True) As Variant
    Dim allErrors As Variant: allErrors = Array()
    Debug.Print "----------- Start DictCollection Compatibility Test ----------------------------"
    
    combineAndPrintTestErrors "Scripting Dictionary Compatibility using Dictionary", testScriptingDictionaryCompatibility(CreateObject("Scripting.Dictionary"), False), allErrors, DebugPrint
    combineAndPrintTestErrors "Scripting Dictionary Compatibility using DictCollection", testScriptingDictionaryCompatibility(New DictCollection, False), allErrors, DebugPrint
    
    combineAndPrintTestErrors "Collection Compatibility using Collection", testCollectionCompatibility(New Collection, False), allErrors, DebugPrint
    combineAndPrintTestErrors "Collection Compatibility using DictCollection", testCollectionCompatibility(New DictCollection, False), allErrors, DebugPrint
    
    Debug.Print "----------- End DictCollection Compatibility Test ------------------------------"
    TestCompatibility = allErrors
End Function

' test compatibility with Scripting.Dictionary behavior (DebugPrint=True -> print errors immediately)
Private Function testScriptingDictionaryCompatibility(d As Object, Optional DebugPrint As Boolean = False) As Variant
    On Error GoTo Fail
    Dim errors As Variant, currentTest As String, i As Long, j As Long, r As Long, removeCount As Long, indexToBeRemoved As Long, indexToBeSwapped1 As Long, indexToBeSwapped2 As Long
    Dim test As Long, Items As Variant, Keys As Variant, testdata As Variant, val1 As Variant, val2 As Variant
    Dim o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, o6 As Object
    errors = Array()
    
    If TypeName(d) = "DictCollection" Then d.EmulateDictionary = True ' switch emulation on if running with DictCollection
    
currentTest = "[DES-1] Dictionary Empty State - Count and CompareMode must be zero"
    If d.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    If d.CompareMode <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[DES-2] Dictionary Empty State - .Items and .Keys must return initialized empty variant array"
    Items = d.Items: Keys = d.Keys
    If VBA.TypeName(Items) <> "Variant()" Or VBA.TypeName(Keys) <> "Variant()" Then
        UtilAddArrayValue errors, ("-> " & currentTest)
    Else
        If UBound(Items) <> -1 Or UBound(Keys) <> -1 Then logError DebugPrint, errors, ("-> " & currentTest)
    End If
    
    For test = 1 To 11
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
                testdata = Array(Array(-100000000000.0001@, "a"), Array(-1.0001@, "b"), Array(CCur(0), "c"), Array(1.0001@, "d"), Array(100000000000.0001@, "e"), Array(922337203685477.5807@, "f"))
            Case 7: ' Decimal
                currentTest = "[DBF-7] Dictionary Basic Functions - Adding, accessing, changing and removing items with keys of datatype Decimal"
                testdata = Array(Array(CDec(-100000000000.0001@), "a"), Array(CDec(-1.0001), "b"), Array(CDec(0), "c"), Array(CDec(1.0001), "d"), Array(CDec(100000000000.0001@), "e"), Array(CDec(922337203685477.5807@), "f"))
            Case 8: ' Date
                currentTest = "[DBF-8] Dictionary Basic Functions - Adding, accessing, changing and removing items with keys of datatype Date"
                testdata = Array(Array(#1/1/100#, "a"), Array(#1/1/1899 10:00:01 AM#, "b"), Array(#1/1/1900#, "c"), Array(#12/31/2000 12:00:00 PM#, "d"), Array(#5/5/2100 11:00:00 PM#, "e"), Array(CDate(#12/31/9999 11:59:59 PM#), "f"))
            Case 9: ' Object or Nothing
                currentTest = "[DBF-9] Dictionary Basic Functions - Adding, accessing, changing and removing items with keys of datatype Object"
                Set o1 = New DictCollection: Set o2 = New DictCollection: Set o3 = New DictCollection: Set o4 = New DictCollection: Set o5 = New DictCollection: Set o6 = New DictCollection
                o1.DefaultValueEnabled = False: o2.DefaultValueEnabled = False: o3.DefaultValueEnabled = False: o4.DefaultValueEnabled = False: o5.DefaultValueEnabled = False: o6.DefaultValueEnabled = False
                testdata = Array(Array(o1, "a"), Array(o2, "b"), Array(o3, "c"), Array(o4, "d"), Array(o5, "e"), Array(o6, "f"), Array(Nothing, "g"))
            Case 10: ' Error
                currentTest = "[DBF-10] Dictionary Basic Functions - Adding, accessing, changing and removing items with keys of datatype Error"
                testdata = Array(Array(CVErr(1), "a"), Array(CVErr(5), "b"), Array(CVErr(11), "c"), Array(CVErr(1000), "d"), Array(CVErr(32000), "e"), Array(CVErr(65535), "f"))
            Case 11: ' Mixed data types
                currentTest = "[DBF-11] Dictionary Basic Functions - Adding, accessing, changing and removing items with mixed key datatypes"
                Set o1 = New DictCollection: Set o2 = New DictCollection: o1.DefaultValueEnabled = False: o2.DefaultValueEnabled = False
                testdata = Array(Array(Null, "null"), Array(Empty, "empty1"), Array(o1, "object1"), Array(CInt(1), "Number=1"), Array("1", "String='1'"), Array(1.00001, "Number=1.00001"), Array(-1.239847E-10, "Number=-1.239847E-10"), Array(o2, "object2"), Array(#12/31/2012 11:12:00 PM#, "Date='12/31/2012 11:12:00 PM'"))
        End Select
        d.RemoveAll
        For i = 0 To UBound(testdata): d.Add testdata(i)(0), testdata(i)(1): Next
        removeCount = UBound(testdata)
        For r = 0 To removeCount
            Items = d.Items: Keys = d.Keys
            If d.Count <> UBound(testdata) + 1 Then logError DebugPrint, errors, ("-> " & currentTest & ": .Count must equal " & UBound(testdata) + 1)
            For i = 0 To UBound(testdata)
                ' compare type of .Item
                val1 = TypeName(d.Item(testdata(i)(0))): val2 = TypeName(testdata(i)(1))
                If val1 <> val2 Then logError DebugPrint, errors, ("-> " & currentTest & ": .Item(" & testdata(i)(0) & ") is " & val1 & " but must be of type " & val2)
                ' compare value of .Item
                If d.Item(testdata(i)(0)) <> testdata(i)(1) Then logError DebugPrint, errors, ("-> " & currentTest & ": .Item(" & testdata(i)(0) & ") must be " & testdata(i)(1))
                ' test .Exists()
                If d.Exists(testdata(i)(0)) <> True Then logError DebugPrint, errors, ("-> " & currentTest & ": .Exists(" & testdata(i)(0) & ") must be true")
                ' compare type of item in .Items array
                val1 = TypeName(Items(i)): val2 = TypeName(testdata(i)(1))
                If val1 <> val2 Then logError DebugPrint, errors, ("-> " & currentTest & ": .Items(" & i & ") is " & val1 & " but must be of type " & val2)
                ' compare value of item in .Items array
                If Items(i) <> testdata(i)(1) Then logError DebugPrint, errors, ("-> " & currentTest & ": .Items(" & i & ") must be " & testdata(i)(1))
                ' compare type of key in .Keys array
                val1 = TypeName(Keys(i)): val2 = TypeName(testdata(i)(0))
                If val1 <> val2 Then logError DebugPrint, errors, ("-> " & currentTest & ": .Keys(" & i & ") is " & val1 & " but must be of type " & val2)
                ' compare value of key in .Keys array
                If IsObject(Keys(i)) And IsObject(testdata(i)(0)) Then
                    If Not (Keys(i) Is testdata(i)(0)) Then logError DebugPrint, errors, ("-> " & currentTest & ": .Keys(" & i & ") must be identical object as original key")
                Else
                    If Keys(i) <> testdata(i)(0) Then logError DebugPrint, errors, ("-> " & currentTest & ": .Keys(" & i & ") must be " & testdata(i)(0))
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
    If val1 <> Empty Then logError DebugPrint, errors, ("-> " & currentTest & ": nonexisting item must be empty")
    If Not d.Exists("a") Then logError DebugPrint, errors, ("-> " & currentTest & ": nonexisting item must be empty")
    If d.Count <> 1 Then logError DebugPrint, errors, ("-> " & currentTest & ": .Count must be different after acessing nonexisting key")
    val1 = d.Item(0)
    If val1 <> Empty Then logError DebugPrint, errors, ("-> " & currentTest & ": nonexisting item must be empty")
    If Not d.Exists(0) Then logError DebugPrint, errors, ("-> " & currentTest & ": nonexisting item must be empty")
    If d.Count <> 2 Then logError DebugPrint, errors, ("-> " & currentTest & ": .Count must be different after acessing nonexisting key")
    Set o1 = New DictCollection: o1.DefaultValueEnabled = False
    val1 = d.Item(o1)
    If val1 <> Empty Then logError DebugPrint, errors, ("-> " & currentTest & ": nonexisting item must be empty")
    If Not d.Exists(o1) Then logError DebugPrint, errors, ("-> " & currentTest & ": nonexisting item must be empty")
    If d.Count <> 3 Then logError DebugPrint, errors, ("-> " & currentTest & ": .Count must be different after acessing nonexisting key")
    Items = d.Items
    If Items(0) <> Empty Or Items(1) <> Empty Or Items(2) <> Empty Then logError DebugPrint, errors, ("-> " & currentTest & ": implicitly added items must be empty!")
    Keys = d.Keys
    If TypeName(Keys(0)) <> "String" Or TypeName(Keys(1)) <> "Integer" Or Not IsObject(Keys(2)) Then
        UtilAddArrayValue errors, ("-> " & currentTest & ": implicitly added keys must have correct types!")
    Else
        If Keys(0) <> "a" Or Keys(1) <> 0 Or Not (Keys(2) Is o1) Then logError DebugPrint, errors, ("-> " & currentTest & ": implicitly added keys must must have correct values!")
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
    If d.Exists("f") Then logError DebugPrint, errors, ("-> " & currentTest & ": key 'f' must not exist after swapping")
    If d.Count <> 5 Then logError DebugPrint, errors, ("-> " & currentTest & ": .Count must be 5 after swapping")
    ' removing item 0, 4, seting "d" to Integer
    d.Remove ("a"): d.Remove ("e"): d.Item("d") = 100
    If (Not IsObject(d.Item("b"))) Or (Not IsObject(d.Item("c"))) Or (Not IsNumeric(d.Item("d"))) Then
        UtilAddArrayValue errors, ("-> " & currentTest & ": wrong item datatypes after remove")
    ElseIf (Not d.Item("b") Is o1) Or (Not d.Item("c") Is o3) Or d.Item("d") <> 100 Then
        UtilAddArrayValue errors, ("-> " & currentTest & ": wrong item values after remove")
    End If
    If d.Exists("a") Or d.Exists("e") Then logError DebugPrint, errors, ("-> " & currentTest & ": key 'a' and 'e' must not exist after swapping")
    If d.Count <> 3 Then logError DebugPrint, errors, ("-> " & currentTest & ": .Count must be 3 after removing")
    
currentTest = "[DCM-1] Dictionary CompareMode - Adding items with binary compare mode"
    d.RemoveAll: d.CompareMode = VbCompareMethod.vbBinaryCompare
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
    d.Add "a‡·‚„‰Â", "LOWER_ASCII_A_CHARS" 'lower case variants of the ASCII character a
    d.Add "A¿¡¬√ƒ≈", "UPPER_ASCII_A_CHARS" 'upper case variants of the ASCII character a
    d.Add ChrW(257) & ChrW(259) & ChrW(261), "LOWER_UNICODE_A_CHARS"  'lower case variants of the UNICODE character a
    d.Add ChrW(256) & ChrW(258) & ChrW(260), "UPPER_UNICODE_A_CHARS"  'upper case variants of the UNICODE character a
currentTest = "[DCM-2] Dictionary CompareMode - Retrieving items with binary compare mode"
    If d("keya") <> "itema" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d("keyA") <> "itemA" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d("keyB") <> "itemB" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d("KeYB") <> "" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d("KeYB") <> Empty Or d("KEYB") <> Empty Or d("keyb") <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
    If d("KEYC") <> "ITEMC" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d("KEyC") <> Empty Or d("Keyc") <> Empty Or d("keyc") <> Empty Then logError DebugPrint, errors, ("-> " & currentTest)
    If d(Chr(128)) <> "EURO128" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d(Chr(164)) <> "EURO164" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d(ChrW(8352)) <> "EURO8352" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d(Chr(128)) <> "EURO128" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d("Ä") <> "EURO128" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d(ChrW(183)) <> "BULLET1" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d(ChrW(8226)) <> "BULLET2" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d(ChrW(8729)) <> "BULLET3" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d(ChrW(9679)) <> "BULLET4" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d(ChrW(9702)) <> "BULLET5" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d("a‡·‚„‰Â") <> "LOWER_ASCII_A_CHARS" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d("A¿¡¬√ƒ≈") <> "UPPER_ASCII_A_CHARS" Then logError DebugPrint, errors, ("-> " & currentTest) ' upper case keys will be sorted before lower case keys
    If d(ChrW(257) & ChrW(259) & ChrW(261)) <> "LOWER_UNICODE_A_CHARS" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d(ChrW(256) & ChrW(258) & ChrW(260)) <> "UPPER_UNICODE_A_CHARS" Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[DCM-3] Dictionary CompareMode - Change of compare mode must not be possible if dictionary contains keys"
    On Error Resume Next: Err.Clear: d.CompareMode = VbCompareMethod.vbTextCompare ' throws error, in DictCollection implementation this is possible
    If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest & ": Change of .CompareMode must throw error 5")
    Err.Clear
    If d.CompareMode <> VbCompareMethod.vbBinaryCompare Then logError DebugPrint, errors, ("-> " & currentTest)
    If d("keya") <> "itema" Then logError DebugPrint, errors, ("-> " & currentTest) ' <- In DictCollection implementation this would now be "itemA"
    If d("keyA") <> "itemA" Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: On Error GoTo Fail
currentTest = "[DCM-4] Dictionary CompareMode - Adding same keys lower/upper case with text compare must not be allowed"
    d.RemoveAll: d.CompareMode = VbCompareMethod.vbTextCompare
    On Error Resume Next
    Err.Clear: d.Add "keya", 1: d.Add "KEYA", 1: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & ": Adding ""keya"" twice must throw error 457")
    Err.Clear: d.Add "a‡·‚„‰Â", 2: d.Add "A¿¡¬√ƒ≈", 2: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & ": Adding ""a‡·‚„‰Â"" twice must throw error 457")
    Err.Clear: d.Add Chr(128), 3: d.Add "Ä", 3: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & ": Adding ""Ä"" twice must throw error 457")
    Err.Clear: On Error GoTo Fail
currentTest = "[DCM-5] Dictionary CompareMode - Change of CompareMode with empty Dictionary must work"
    d.RemoveAll
    d.CompareMode = VbCompareMethod.vbBinaryCompare: If d.CompareMode <> VbCompareMethod.vbBinaryCompare Then logError DebugPrint, errors, ("-> " & currentTest)
    d.CompareMode = VbCompareMethod.vbTextCompare: If d.CompareMode <> VbCompareMethod.vbTextCompare Then logError DebugPrint, errors, ("-> " & currentTest)
    d.CompareMode = VbCompareMethod.vbDatabaseCompare: If d.CompareMode <> VbCompareMethod.vbDatabaseCompare Then logError DebugPrint, errors, ("-> " & currentTest)
    d.CompareMode = VbCompareMethod.vbBinaryCompare ' set back to default
    
currentTest = "[DKC-1] Dictionary Key Changing - Changing the key of existing items"
    d.RemoveAll: d.Add "a", "itemD": d.Add "b", "itemE": d.Add "c", "itemF": d.key("a") = "d": d.key("b") = "e": d.key("c") = "f"
    ' STRANGE: d.Item("d") <> "itemD" returns true
    If (StrComp(d.Item("d"), "itemD") <> 0) Or (StrComp(d.Item("e"), "itemE") <> 0) Or (StrComp(d.Item("f"), "itemF") <> 0) Then logError DebugPrint, errors, ("-> " & currentTest & ": string keys")
    d.Add -1, "itemD": d.Add 0, "itemE": d.Add 1, "itemF": d.key(-1) = -100: d.key(0) = 2: d.key(1) = 100
    If (StrComp(d.Item(-100), "itemD") <> 0) Or (StrComp(d.Item(2), "itemE") <> 0) Or (StrComp(d.Item(100), "itemF") <> 0) Then logError DebugPrint, errors, ("-> " & currentTest & ": number keys")
    ' changing object keys
    Set o1 = New DictCollection: Set o2 = New DictCollection: Set o3 = New DictCollection: Set o4 = New DictCollection: Set o5 = New DictCollection: Set o6 = New DictCollection
    o1.DefaultValueEnabled = False: o2.DefaultValueEnabled = False: o3.DefaultValueEnabled = False: o4.DefaultValueEnabled = False: o5.DefaultValueEnabled = False: o6.DefaultValueEnabled = False
    d.Add o1, "itemD": d.Add o2, "itemE": d.Add o3, "itemF"
    d.key(o1) = o4: d.key(o2) = o5: d.key(o3) = o6
    If (StrComp(d.Item(o4), "itemD") <> 0) Or (StrComp(d.Item(o5), "itemE") <> 0) Or (StrComp(d.Item(o6), "itemF") <> 0) Then logError DebugPrint, errors, ("-> " & currentTest & ": object keys")
currentTest = "[DKC-2] Dictionary Key Changing - Changing the key of existing items from lower to upper case must work with .CompareMethod = 1 (case insensitive)"
    d.RemoveAll: d.CompareMode = VbCompareMethod.vbTextCompare
    d.Add "a", "ItemA": d.key("a") = "A"
    If Not d.Exists("A") Then logError DebugPrint, errors, ("-> " & currentTest)
    d.RemoveAll: d.CompareMode = VbCompareMethod.vbBinaryCompare ' set back to default
currentTest = "[DKC-3] Dictionary Key Changing - Assigning keys already in use must throw error 457"
    Set o1 = New DictCollection: Set o2 = New DictCollection: o1.DefaultValueEnabled = False: o2.DefaultValueEnabled = False: o1.Add "o1", "o1": o2.Add "o2", "o2"
    d.RemoveAll: d.Add "a", "itemA": d.Add "b", "itemB": d.Add 1, "itemC": d.Add 2, "itemD": d.Add o1, "itemE": d.Add o2, "itemF"
    On Error Resume Next
    Err.Clear: d.key("a") = "b": If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: d.key("b") = "a": If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: d.key(1) = 2: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: d.key(2) = "b": If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: d.key(o1) = o2: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: d.key(o2) = "a": If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: On Error GoTo Fail
currentTest = "[DKC-3] Dictionary Key Changing - Changing number keys to same number but with different datatype should not have any effect"
    testdata = Array(Array(CInt(1), "Integer"), Array(CLng(1), "Long"), Array(CSng(1), "Single"), Array(CDbl(1), "Double"), Array(CCur(1), "Currency"), Array(CDec(1), "Decimal"), Array(#12/31/1899#, "Date"))
    For i = 0 To UBound(testdata)
        d.RemoveAll: d.Add testdata(i)(0), testdata(i)(1)
        For j = 0 To UBound(testdata)
            If j <> i Then
                d.key(testdata(i)(0)) = testdata(j)(0): Keys = d.Keys
                If TypeName(Keys(0)) <> TypeName(testdata(i)(0)) Then
                    UtilAddArrayValue errors, ("-> " & currentTest & ": re-assignment with " & TypeName(testdata(j)(0)) & " changed original key datatype of " & TypeName(testdata(i)(0)))
                End If
            End If
        Next
    Next

currentTest = "[DNIA-1] Dictionary Nonexisting item access - Nonexisting keys must return empty variant"
    d.RemoveAll
    val1 = d.Item("nonexisting"): If Not IsEmpty(val1) Then logError DebugPrint, errors, ("-> " & currentTest)
    val1 = d.Item(0): If Not IsEmpty(val1) Then logError DebugPrint, errors, ("-> " & currentTest)
    
currentTest = "[DAE-1] Dictionary Access Errors - .Add without arguments must throw Error 450"
    On Error Resume Next
    Err.Clear: d.RemoveAll:  d.Add: If Err.Number <> 450 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: On Error GoTo Fail
    
currentTest = "[DAE-2] Dictionary Access Errors - Adding existing key must throw error 457"
    On Error Resume Next
    d.RemoveAll: Err.Clear: d.Add "1", 1: If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: String).")
    d.Add "1", 2: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & " (Keytype: String)")
    Err.Clear: d.Add CInt(1), 1: If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: Integer).")
    Err.Clear: d.Add CInt(1), 2: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & " (Keytype: Integer)")
    Err.Clear: d.Add CLng(2), 1: If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: Long).")
    Err.Clear: d.Add CLng(2), 2: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & " (Keytype: Long)")
    Err.Clear: d.Add CSng(3), 1: If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: Single).")
    Err.Clear: d.Add CSng(3), 2: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & " (Keytype: Single)")
    Err.Clear: d.Add CDbl(4), 1: If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: Double).")
    Err.Clear: d.Add CDbl(4), 2: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & " (Keytype: Double)")
    Err.Clear: d.Add CCur(5), 1: If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: Currency).")
    Err.Clear: d.Add CCur(5), 2: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & " (Keytype: Currency)")
    Err.Clear: d.Add CDec(6), 1: If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: Decimal).")
    Err.Clear: d.Add CDec(6), 2: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & " (Keytype: Decimal)")
    Err.Clear: d.Add CDate(7), 1: If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: Date).")
    Err.Clear: d.Add CDate(7), 2: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & " (Keytype: Date)")
    Err.Clear: d.Add CBool(1), 1: If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: Boolean).")
    Err.Clear: d.Add CBool(1), 2: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & " (Keytype: Boolean)")
    Err.Clear: d.Add Empty, 1: If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: Empty).")
    Err.Clear: d.Add Empty, 2: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & " (Keytype: Empty)")
    d.RemoveAll
    Err.Clear: d.Add "", 1: If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: String '').")
    Err.Clear: d.Add Empty, 2: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & " (Keytype: Empty with existing String '' )")
    d.RemoveAll
    Err.Clear: d.Add Empty, 1: If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: Empty).")
    Err.Clear: d.Add "", 2: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & " (Keytype: String '' with existing Empty)")
    d.RemoveAll
    Err.Clear: d.Add "", 1: If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: String '').")
    Err.Clear: d.Add "", 2: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & " (Keytype: String '' with existing String '' )")
    Set o1 = New DictCollection: Err.Clear: d.Add o1, 1: If Err.Number <> 0 Then logError DebugPrint, errors, ("-> " & currentTest & ": Adding first nonexisting key failed (Keytype: Object).")
    d.Add o1, 2: If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest & " (Keytype: Object)")
    Err.Clear: On Error GoTo Fail
    
currentTest = "[DAE-3] Dictionary Access Errors - Changing nonexisting key must throw Error 32811"
    On Error Resume Next
    d.RemoveAll: Err.Clear: d.key("b") = "c": If Err.Number <> 32811 Then logError DebugPrint, errors, ("-> " & currentTest & " (Keytype: String)")
    Err.Clear: d.key(CInt(1)) = 2: If Err.Number <> 32811 Then logError DebugPrint, errors, ("-> " & currentTest & " (Keytype: Integer)")
    Err.Clear: d.Add 1, "a": d.key(2) = 2: If Err.Number <> 32811 Then logError DebugPrint, errors, ("-> " & currentTest & " (Keytype: Integer)")
    Err.Clear: Set o1 = New DictCollection: o1.DefaultValueEnabled = False: d.key(o1) = 2: If Err.Number <> 32811 Then logError DebugPrint, errors, ("-> " & currentTest & " (Keytype: Object)")
    Err.Clear: On Error GoTo Fail
    
    GoTo ExitFunction
Fail:
    logError DebugPrint, errors, ("-> " & currentTest & " -> Error " & Err.Number & " " & Err.Description)
ExitFunction:
    testScriptingDictionaryCompatibility = errors
End Function

' test compatibility with Collection behavior (DebugPrint=True -> print errors immediately)
Private Function testCollectionCompatibility(d As Object, Optional DebugPrint As Boolean = False) As Variant
    On Error GoTo Fail
    Dim errors As Variant, currentTest As String, i As Long, j As Long, r As Long, key As String, removeCount As Long, indexToBeRemoved As Long, indexToBeSwapped1 As Long, indexToBeSwapped2 As Long
    Dim test As Long, Items As Variant, Keys As Variant, testdata As Variant, val1 As Variant, val2 As Variant
    Dim o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, o6 As Object, isDC As Boolean
    errors = Array()
    
    ' set emulation
    Select Case TypeName(d)
        Case "DictCollection": d.EmulateCollection = True: isDC = True
        Case Else: isDC = False
    End Select

currentTest = "[CES-1] Collection Empty State - Count must be zero"
    If d.Count <> 0 Then logError DebugPrint, errors, ("-> " & currentTest)
    
    For i = d.Count To 1 Step -1: d.Remove i: Next  ' remove all
    For test = 1 To 11
        Select Case test
            Case 1: ' String
                currentTest = "[CBF-1] Collection Basic Functions - Adding, accessing, changing and removing items of datatype String"
                testdata = Array(Array("", "a"), Array(UtilGetMissingValue, "b"), Array("keyC", "c"), Array("keyd", "d"), Array("KEYE", "e"), Array("kEYf", "f"))
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
                testdata = Array(Array("a", -100000000000.0001@), Array("b", -1.0001@), Array("c", CCur(0)), Array("d", 1.0001@), Array("e", 100000000000.0001@), Array("f", 922337203685477.5807@))
            Case 7: ' Decimal
                currentTest = "[CBF-7] Collection Basic Functions - Adding, accessing, changing and removing items of datatype Decimal"
                testdata = Array(Array("a", CDec(-100000000000.0001@)), Array("b", CDec(-1.0001)), Array("c", CDec(0)), Array("d", CDec(1.0001)), Array("e", CDec(100000000000.0001@)), Array("f", CDec(922337203685477.5807@)))
            Case 8: ' Date
                currentTest = "[CBF-8] Collection Basic Functions - Adding, accessing, changing and removing items of datatype Date"
                testdata = Array(Array("a", #1/1/100#), Array("b", #1/1/1899 10:00:01 AM#), Array("c", #1/1/1900#), Array("d", #12/31/2000 12:00:00 PM#), Array("e", #5/5/2100 11:00:00 PM#), Array("f", CDate(#12/31/9999 11:59:59 PM#)))
            Case 9: ' Object
                currentTest = "[CBF-9] Collection Basic Functions - Adding, accessing, changing and removing items of datatype Object"
                Set o1 = New DictCollection: Set o2 = New DictCollection: Set o3 = New DictCollection: Set o4 = New DictCollection: Set o5 = New DictCollection: Set o6 = New DictCollection
                o1.DefaultValueEnabled = False: o2.DefaultValueEnabled = False: o3.DefaultValueEnabled = False: o4.DefaultValueEnabled = False: o5.DefaultValueEnabled = False: o6.DefaultValueEnabled = False
                testdata = Array(Array("a", o1), Array("b", o2), Array("c", o3), Array("d", o4), Array("e", o5), Array("f", o6))
            Case 10: ' Error
                currentTest = "[CBF-10] Collection Basic Functions - Adding, accessing, changing and removing items of datatype Error"
                testdata = Array(Array("a", CVErr(1)), Array("b", CVErr(5)), Array("c", CVErr(11)), Array("d", CVErr(1000)), Array("e", CVErr(32000)), Array("f", CVErr(65535)))
            Case 11: ' Mixed data types
                currentTest = "[CBF-11] Collection Basic Functions - Adding, accessing, changing and removing items of mixed datatypes"
                Set o1 = New DictCollection: Set o2 = New DictCollection: o1.DefaultValueEnabled = False: o2.DefaultValueEnabled = False
                testdata = Array(Array("", Empty), Array("object1", o1), Array("Number=1", CInt(1)), Array("String1", "1"), Array("Number=1.00001", 1.00001), Array("Number=-1.239847E-10", -1.239847E-10), Array("object2", o2), Array("Date='12/31/2012 11:12:00 PM'", #12/31/2012 11:12:00 PM#))
        End Select
        For i = 0 To UBound(testdata)
             If isDC Then d.Add2 testdata(i)(1), testdata(i)(0) Else d.Add testdata(i)(1), testdata(i)(0)
        Next
        removeCount = UBound(testdata)
        For r = 0 To removeCount
            If d.Count <> UBound(testdata) + 1 Then logError DebugPrint, errors, ("-> " & currentTest & ": .Count must equal " & UBound(testdata) + 1)
            For i = 0 To UBound(testdata)
                ' compare type of .Item using index
                val1 = TypeName(d.Item(i + 1)): val2 = TypeName(testdata(i)(1))
                If val1 <> val2 Then
                    UtilAddArrayValue errors, ("-> " & currentTest & ": Access with index failed - .Item(" & i + 1 & ") is " & val1 & " but must be of type " & val2)
                ElseIf IsObject(d.Item(i + 1)) Then
                    If Not (d.Item(i + 1) Is testdata(i)(1)) Then logError DebugPrint, errors, ("-> " & currentTest & ": Access with index failed - .Item(" & i + 1 & ") must be same object as testdata index=" & i)
                ElseIf d.Item(i + 1) <> testdata(i)(1) Then
                    UtilAddArrayValue errors, ("-> " & currentTest & ": Access with index failed - .Item(" & i + 1 & ") must be " & testdata(i)(1))
                End If
                
                If TypeName(testdata(i)(0)) = "String" Then
                    ' compare type of .Item using key
                    val1 = TypeName(d.Item(testdata(i)(0))): val2 = TypeName(testdata(i)(1))
                    If val1 <> val2 Then
                        UtilAddArrayValue errors, ("-> " & currentTest & ": Access with key failed - .Item(" & testdata(i)(0) & ") is " & val1 & " but must be of type " & val2)
                    ElseIf IsObject(d.Item(testdata(i)(0))) Then
                        If Not (d.Item(testdata(i)(0)) Is testdata(i)(1)) Then logError DebugPrint, errors, ("-> " & currentTest & ": Access with key failed - .Item(" & i + 1 & ") must be same object as testdata index=" & i)
                    ElseIf d.Item(testdata(i)(0)) <> testdata(i)(1) Then
                        UtilAddArrayValue errors, ("-> " & currentTest & ": Access with key failed - .Item(" & i + 1 & ") must be " & testdata(i)(1))
                    End If
                    ' compare access with uppercase key
                    key = UCase(testdata(i)(0))
                    If IsObject(d.Item(key)) Then
                        If Not (d.Item(key) Is testdata(i)(1)) Then logError DebugPrint, errors, ("-> " & currentTest & ": Access with uppercase key failed - .Item(" & key & ") must be " & testdata(i)(1))
                    ElseIf d.Item(key) <> testdata(i)(1) Then
                        UtilAddArrayValue errors, ("-> " & currentTest & ": Access with uppercase key failed - .Item(" & key & ") must be " & testdata(i)(1))
                    End If
                    ' compare access with lowercase key
                    key = LCase(testdata(i)(0))
                    If IsObject(d.Item(key)) Then
                        If Not (d.Item(key) Is testdata(i)(1)) Then logError DebugPrint, errors, ("-> " & currentTest & ": Access with lowercase key failed - .Item(" & key & ") must be " & testdata(i)(1))
                    ElseIf d.Item(key) <> testdata(i)(1) Then
                        UtilAddArrayValue errors, ("-> " & currentTest & ": Access with lowercase key failed - .Item(" & key & ") must be " & testdata(i)(1))
                    End If
                Else
                    ' keys with type other than "String" must throw errors
                    Err.Clear: On Error Resume Next
                    val1 = d.Item(testdata(i)(0)): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest & ": Access with non-string key did not throw error!")
                    On Error GoTo Fail: Err.Clear
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
                        d.Add2 testdata(indexToBeSwapped1)(1), testdata(indexToBeSwapped1)(0), indexToBeSwapped1 + 1 ' add before
                    Else
                        d.Add2 testdata(indexToBeSwapped1)(1), testdata(indexToBeSwapped1)(0), , indexToBeSwapped1 ' add after
                    End If
                    If indexToBeSwapped2 = 0 Then
                        d.Add2 testdata(indexToBeSwapped2)(1), testdata(indexToBeSwapped2)(0), indexToBeSwapped2 + 1 ' add before
                    Else
                        d.Add2 testdata(indexToBeSwapped2)(1), testdata(indexToBeSwapped2)(0), , indexToBeSwapped2 ' add after
                    End If
                Else
                    If indexToBeSwapped1 = 0 Then
                        d.Add testdata(indexToBeSwapped1)(1), testdata(indexToBeSwapped1)(0), indexToBeSwapped1 + 1 ' add before
                    Else
                        d.Add testdata(indexToBeSwapped1)(1), testdata(indexToBeSwapped1)(0), , indexToBeSwapped1 ' add after
                    End If
                    If indexToBeSwapped2 = 0 Then
                        d.Add testdata(indexToBeSwapped2)(1), testdata(indexToBeSwapped2)(0), indexToBeSwapped2 + 1 ' add before
                    Else
                        d.Add testdata(indexToBeSwapped2)(1), testdata(indexToBeSwapped2)(0), , indexToBeSwapped2 ' add after
                    End If
                End If
            End If
            If TypeName(testdata(indexToBeRemoved)(0)) = "String" Then d.Remove testdata(indexToBeRemoved)(0) Else d.Remove indexToBeRemoved + 1
            UtilRemoveArrayValueByIndex testdata, indexToBeRemoved ' remove 1 item in test array
        Next
    Next
    
currentTest = "[CIA-1] Collection Index Access"
    For i = d.Count To 1 Step -1: d.Remove i: Next  ' remove all
    If isDC Then d.Add2 "a": d.Add2 "b" Else d.Add "a": d.Add "b"
    If d.Item(CInt(1)) <> "a" Or d.Item(CInt(2)) <> "b" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d.Item(CLng(1)) <> "a" Or d.Item(CLng(2)) <> "b" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d.Item(CSng(1.4)) <> "a" Or d.Item(CSng(1.6)) <> "b" Or d.Item(CSng(1.5)) <> "b" Or d.Item(CSng(2#)) <> "b" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d.Item(CDbl(1.4)) <> "a" Or d.Item(CDbl(1.6)) <> "b" Or d.Item(CDbl(1.5)) <> "b" Or d.Item(CDbl(2#)) <> "b" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d.Item(CCur(1.4)) <> "a" Or d.Item(CCur(1.6)) <> "b" Or d.Item(CCur(1.5)) <> "b" Or d.Item(CCur(2#)) <> "b" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d.Item(CDec(1.4)) <> "a" Or d.Item(CDec(1.6)) <> "b" Or d.Item(CDec(1.5)) <> "b" Or d.Item(CDec(2#)) <> "b" Then logError DebugPrint, errors, ("-> " & currentTest)
    If d.Item(CDate(1.4)) <> "a" Or d.Item(CDate(1.6)) <> "b" Or d.Item(CDate(1.5)) <> "b" Or d.Item(CDate(2#)) <> "b" Then logError DebugPrint, errors, ("-> " & currentTest)
    
    
currentTest = "[CCIA-1] Collection Case Insensitive Access"
    For i = d.Count To 1 Step -1: d.Remove i: Next  ' remove all
    If isDC Then
        d.Add2 "ASCII_A_CHARS", "a‡·‚„‰Â" 'lower case variants of the ASCII character a
    Else
        d.Add "ASCII_A_CHARS", "a‡·‚„‰Â" 'lower case variants of the ASCII character a
    End If
    If d.Item("A¿¡¬√ƒ≈") <> "ASCII_A_CHARS" Then logError DebugPrint, errors, ("-> " & currentTest)
    If isDC Then
        d.Add2 "UNICODE_A_CHARS", ChrW(257) & ChrW(259) & ChrW(261) 'lower case variants of the ASCII character a
    Else
        d.Add "UNICODE_A_CHARS", ChrW(257) & ChrW(259) & ChrW(261) 'lower case variants of the ASCII character a
    End If
    If d.Item(ChrW(256) & ChrW(258) & ChrW(260)) <> "UNICODE_A_CHARS" Then logError DebugPrint, errors, ("-> " & currentTest)

currentTest = "[CAE-1] Collection Access Errors - Accessing nonexisting keys in empty Collection must throw Error 5 'Invalid procedure call or argument'"
    For i = d.Count To 1 Step -1: d.Remove i: Next  ' remove all
    On Error Resume Next
    Err.Clear: val1 = d.Item(Empty): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item(""): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item("a"): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[CAE-2] Collection Access Errors - Accessing nonexisting indexes in empty Collection must throw Error 5 'Invalid procedure call or argument'"
    Err.Clear: val1 = d.Item(0): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item(1): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item(-1): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item(1000): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[CAE-3] Collection Access Errors - Removing nonexisting indexes in empty Collection must throw Error 5 'Invalid procedure call or argument'"
    Err.Clear: val1 = d.Remove(0): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Remove(1): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Remove(-1): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Remove(1000): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[CAE-4] Collection Access Errors - Removing nonexisting keys must in empty Collection throw Error 5 'Invalid procedure call or argument'"
    Err.Clear: val1 = d.Remove(Empty): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Remove(""): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Remove("a"): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[CAE-5] Collection Access Errors - Accessing invalid keytypes in empty Collection must throw Error 5 'Invalid procedure call or argument'"
    Err.Clear: val1 = d.Remove(DictCollection.UtilGetMissingValue): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item(Null): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item(Nothing): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item(CVErr(11)): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[CAE-6] Collection Access Errors - Accessing nonexisting keys in filled Collection must throw Error 9 'Subscript out of range'"
    For i = d.Count To 1 Step -1: d.Remove i: Next  ' remove all
    If isDC Then d.Add2 "a" Else d.Add "a"
    Err.Clear: val1 = d.Item(Empty): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[CAE-7] Collection Access Errors - Accessing nonexisting String keys in filled Collection must throw Error 5 'Invalid procedure call or argument'"
    Err.Clear: val1 = d.Item(""): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item("b"): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[CAE-8] Collection Access Errors - Accessing nonexisting indexes in filled Collection must throw Error 9 'Subscript out of range'"
    Err.Clear: val1 = d.Item(0): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item(2): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item(-1): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item(1000): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[CAE-9] Collection Access Errors - Removing nonexisting indexes in filled Collection must throw Error 9 'Subscript out of range'"
    Err.Clear: val1 = d.Remove(0): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Remove(3): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Remove(-1): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Remove(1000): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[CAE-10] Collection Access Errors - Removing nonexisting keys in filled Collection must throw Error 5 'Invalid procedure call or argument'"
    Err.Clear: val1 = d.Remove(""): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Remove("a"): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[CAE-11] Collection Access Errors - Removing key Empty in filled Collection must throw Error 9 'Subscript out of range'"
    Err.Clear: val1 = d.Remove(Empty): If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[CAE-12] Collection Access Errors - Removing key of type Missing in filled Collection must throw Error 5 'Invalid procedure call or argument'"
    Err.Clear: val1 = d.Remove(DictCollection.UtilGetMissingValue): If Err.Number <> 5 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[CAE-13] Collection Access Errors - Accessing keys of invalid datatype in filled Collection must throw Error 13 'Type mismatch'"
    Err.Clear: val1 = d.Item(Null): If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item(Nothing): If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: val1 = d.Item(CVErr(11)): If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[CAE-14] Collection Access Errors - Adding existing keys must throw Error 457 'This key is already associated with an element of this collection'"
    For i = d.Count To 1 Step -1: d.Remove i: Next  ' remove all
    If isDC Then d.Add2 "itemA", "keyA": d.Add2 "itemB", "" Else d.Add "itemA", "keyA": d.Add "itemB", ""
    Err.Clear: If isDC Then d.Add2 "itemC", "keya" Else d.Add "itemC", "keya"
        If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: If isDC Then d.Add2 "itemC", "" Else d.Add "itemC", ""
        If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[CAE-15] Collection Access Errors - Adding keys of invalid datatype must throw Error 13 'Type mismatch'"
    For i = d.Count To 1 Step -1: d.Remove i: Next  ' remove all
    testdata = Array(CInt(1), CLng(1), CSng(1), CDbl(1), CCur(1), CDec(1), CDate(1), Null, Nothing)
    For i = 0 To UBound(testdata)
        Err.Clear
        If isDC Then d.Add2 "item", testdata(i) Else d.Add "item", testdata(i)
        If Err.Number <> 13 Then logError DebugPrint, errors, ("-> " & currentTest)
    Next
currentTest = "[CAE-16] Collection Access Errors - Inserting items with before or after invalid index must throw Error 9 'Subscript out of range'"
    For i = d.Count To 1 Step -1: d.Remove i: Next  ' remove all
    If isDC Then d.Add2 "itemA": d.Add2 "itemB" Else d.Add "itemA": d.Add "itemB" ' index 1,2
    Err.Clear: If isDC Then d.Add2 "itemC", , 0 Else d.Add "itemC", , 0 ' before -1
        If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: If isDC Then d.Add2 "itemC", , 0 Else d.Add "itemC", , 0 ' before 0
        If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: If isDC Then d.Add2 "itemC", , 3 Else d.Add "itemC", , 3 ' before 3
        If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: If isDC Then d.Add2 "itemC", , , 0 Else d.Add "itemC", , 0 ' after 0
        If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: If isDC Then d.Add2 "itemC", , , 3 Else d.Add "itemC", , 3 ' after 3
        If Err.Number <> 9 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[CAE-17] Collection Access Errors - Adding items with existing indexes as keys must throw Error 457 'This key is already associated with an element of this collection'"
    For i = d.Count To 1 Step -1: d.Remove i: Next  ' remove all
    If isDC Then d.Add2 "itemA": d.Add2 "itemB" Else d.Add "itemA": d.Add "itemB" ' index 1,2
    Err.Clear: If isDC Then d.Add2 "itemC", 1 Else d.Add "itemC", 1 ' index=1
        If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest)
    Err.Clear: If isDC Then d.Add2 "itemC", 2 Else d.Add "itemC", 2 ' index=2
        If Err.Number <> 457 Then logError DebugPrint, errors, ("-> " & currentTest)
currentTest = "[CAE-18] Collection Access Errors - Adding items using invalid datatye for Before or After must throw Error 5, 9 or 13"
    For i = d.Count To 1 Step -1: d.Remove i: Next  ' remove all
    Set o1 = New DictCollection: o1.DefaultValueEnabled = False
    testdata = Array(Array("a", 5), Array(Empty, 9), Array(CVErr(1), 13), Array(o1, 13), Array(Nothing, 13), Array(Null, 13))
    If isDC Then d.Add2 "itemA": d.Add2 "itemB" Else d.Add "itemA": d.Add "itemB" ' index 1,2
    For i = 0 To UBound(testdata)
        Err.Clear: If isDC Then d.Add2 "itemC", , testdata(i)(0) Else d.Add "itemC", , testdata(i)(0) ' before = invalid datatype
        If Err.Number <> testdata(i)(1) Then logError DebugPrint, errors, ("-> " & currentTest)
        Err.Clear: If isDC Then d.Add2 "itemC", , , testdata(i)(0) Else d.Add "itemC", , , testdata(i)(0) ' after = invalid datatype
        If Err.Number <> testdata(i)(1) Then logError DebugPrint, errors, ("-> " & currentTest)
    Next
    
    Err.Clear: On Error GoTo Fail

' Adding items with Empty or Error (CVErr(5)) as key throws Error 13 'Type Mismatch'
' Adding items with Missing as key is OK but accessing them with Missing throws Error 5 'Invalid procedure call or argument'

    GoTo ExitFunction
Fail:
    logError DebugPrint, errors, ("-> " & currentTest & " -> Error " & Err.Number & " " & Err.Description)
ExitFunction:
    testCollectionCompatibility = errors
End Function
' ======================== END: COMPATIBILITY TESTS ===============================================================================================


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
' - Arrays cannot be sized using the To keyword: 'Dim Arr(first To last)' does not work
' - CVErr() cannot be used as in VBA, throws error
' - IsMissing() does not exist, has to be implemented as (VarType(p) = vbError)
' - iif(cond,truepart,falsepart) does not exist, has to implemented as if ... then ... else
' - implicit casts have to be made explicit using CInt(), CLng(), etc.
' - Class default property is declared as 'Public Default Property Get MyProperty()'
' - File functions do not work: SetAtt, Kill, FileLen ...
' ======================== END: VBSCRIPT INFO =====================================================================================================

