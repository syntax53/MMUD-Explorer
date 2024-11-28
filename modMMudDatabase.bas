Attribute VB_Name = "modMMudDatabase"
Option Explicit
Option Base 0

Global UseExpMulti As Boolean

Public nSpellNest As Integer

Public DB As Database
'Public WS As Workspace
Public tabItems As Recordset
Public tabClasses As Recordset
Public tabRaces As Recordset
Public tabSpells As Recordset
Public tabInfo As Recordset
Public tabMonsters As Recordset
Public tabShops As Recordset
Public tabRooms As Recordset
Public tabTBInfo As Recordset
Public tabTempRS As Recordset

Public nAveragePossSpawns As Currency
Public nAverageMobsPerLair As Currency
Public nMonsterDamage() As Currency
Public nMonsterPossy() As Currency
Public nMonsterSpawnChance() As Currency
Public nMonsterAVGLairExp() As Currency
Public bQuickSpell As Boolean

Public Type SpellMinMaxDur
    nMin As Currency
    nMax As Currency
    nDur As Currency
    sMin As String
    sMax As String
    sDur As String
    bNoHeader As Boolean
End Type

Public Type typItemCostDetail
    Cost As Long
    Coin As Long
End Type

Public Enum enmMagicEnum
    None = 0
    Mage = 1
    Priest = 2
    Druid = 3
    Bard = 4
    Kai = 5
End Enum

Public Enum MVDatType
    item = 0
    Spell = 1
    Monster = 2
    Shop = 3
    Class = 4
    Race = 5
    Text = 6
    Room = 7
End Enum

Public Function CalcExpNeededByRaceClass(ByVal nLevel As Long, ByVal nClass As Long, ByVal nRace As Long) As Currency
Dim nClassExp As Integer, nRaceExp As Integer, nExp As Currency, nChart As Long

On Error GoTo error:

If nClass > 0 Then
    tabClasses.Index = "pkClasses"
    tabClasses.Seek "=", nClass
    If tabClasses.NoMatch = True Then
        nClassExp = 0
    Else
        nClassExp = tabClasses.Fields("ExpTable") + 100
    End If
End If

If nRace > 0 Then
    tabRaces.Index = "pkRaces"
    tabRaces.Seek "=", nRace
    If tabRaces.NoMatch = True Then
        nRaceExp = 0
    Else
        nRaceExp = tabRaces.Fields("ExpTable")
    End If
End If

nChart = nClassExp + nRaceExp
nExp = CalcExpNeeded(nLevel, nChart)
CalcExpNeededByRaceClass = Fix(nExp * 10000)

Exit Function
error:
Call HandleError("CalcExpNeededByRaceClass")

End Function
Public Function OpenTables(sFile As String) As Boolean
On Error GoTo error:

UseExpMulti = False

'Set WS = DAO.CreateWorkspace("MMUD_Explorer_WS", "MMUD_Explorer", False, dbUseJet)
Set DB = OpenDatabase(sFile, False, True)

sCurrentDatabaseFile = sFile

Set tabItems = DB.OpenRecordset("Items")
Set tabClasses = DB.OpenRecordset("Classes")
Set tabRaces = DB.OpenRecordset("Races")
Set tabSpells = DB.OpenRecordset("Spells")
Set tabMonsters = DB.OpenRecordset("Monsters")
Set tabShops = DB.OpenRecordset("Shops")
Set tabRooms = DB.OpenRecordset("Rooms")
Set tabInfo = DB.OpenRecordset("Info")
Set tabTBInfo = DB.OpenRecordset("TBInfo")

Call TestMonExpMulti

OpenTables = True

Exit Function
error:
Call HandleError("OpenDatabase")
'Resume Next
End Function

Public Sub CalculateAverageLairs()
Dim nTotalLairs As Long, nTotalTOTALLairs As Long, nLairs As Long, nTotalPossy As Currency, nSpawnChance As Currency
Dim sPattern As String, sArr() As String, sArr2() As String, x As Integer, nMapRoom As RoomExitType, nPos As Integer
Dim dictLairs As Object, sRoomKey As String, nMobsTotal As Currency, nMobs As Currency, sLairValuePattern As String
Dim sValues(2) As String
On Error GoTo error:

Set tabTempRS = DB.OpenRecordset( _
    "SELECT [Number],[Summoned By] FROM Monsters WHERE [Summoned By] Like ""*(lair)*""", dbOpenSnapshot)

sPattern = "Group\(lair\): \d+\/\d+"
If nNMRVer >= 1.82 Then sPattern = "\[\d+\]" & sPattern
If nNMRVer >= 1.83 Then sPattern = "\[\d+\]\[\d+\]" & sPattern

sLairValuePattern = "^\[(\d+)\]\[(\d+)\]\[(\d+)\]"

If Not tabTempRS.EOF Then
    tabTempRS.MoveFirst
    ' Create a Dictionary object
    Set dictLairs = CreateObject("Scripting.Dictionary")
    dictLairs.CompareMode = vbTextCompare ' Case-insensitive comparison

    Do While Not tabTempRS.EOF
        nMobs = 0
        nMobsTotal = 0
        nLairs = 0
        sArr() = RegExpFind(tabTempRS.Fields("Summoned By"), sPattern)
        If UBound(sArr()) >= 0 Then
            nLairs = UBound(sArr()) + 1
            nTotalTOTALLairs = nTotalTOTALLairs + nLairs
            nSpawnChance = 0
            For x = 0 To UBound(sArr())
                '[1]Group(lair): 1/1239
                If nNMRVer >= 1.83 Then
                    sValues(0) = 0 'average exp for lair
                    sValues(1) = 0 'unique mobs that could spawn
                    sValues(2) = 0 'max regen of lair
                    sArr2() = RegExpFind(sArr(x), sLairValuePattern, , , 3)
                    If UBound(sArr2()) >= 0 Then sValues(0) = sArr2(0)
                    If UBound(sArr2()) >= 1 Then sValues(1) = sArr2(1)
                    If UBound(sArr2()) >= 2 Then sValues(2) = sArr2(2)
                    
                    nMonsterAVGLairExp(tabTempRS.Fields("Number")) = nMonsterAVGLairExp(tabTempRS.Fields("Number")) + Val(sValues(0))
                    nMobs = Val(sValues(2))
                    
                    If Val(sValues(1)) > 0 And nMobs > 0 Then
                        nSpawnChance = nSpawnChance + Round(1 - (1 - (1 / Val(sValues(1)))) ^ nMobs, 2)
                    End If
                ElseIf nNMRVer >= 1.82 Then
                    nMobs = ExtractNumbersFromString(sArr(x))
                    nSpawnChance = nSpawnChance + 1
                Else
                    nMobs = 1
                    nSpawnChance = nSpawnChance + 1
                End If
                nMobsTotal = nMobsTotal + nMobs
                
                nPos = InStr(1, sArr(x), "Group(lair): ", vbTextCompare)
                If nPos > 0 Then
                    sRoomKey = "[" & Mid(sArr(x), nPos + Len("Group(lair): "), Len(sArr(x)) - nPos - Len("Group(lair): ") + 1) & "]"
                    If Not dictLairs.Exists(sRoomKey) Then
                        dictLairs.Add sRoomKey, True
                        nTotalLairs = nTotalLairs + 1
                        nTotalPossy = nTotalPossy + nMobs
                    End If
                End If
            Next x
            
            nMonsterAVGLairExp(tabTempRS.Fields("Number")) = Round(nMonsterAVGLairExp(tabTempRS.Fields("Number")) / nLairs) 'average exp per lair for mob
            nMonsterPossy(tabTempRS.Fields("Number")) = Round(nMobsTotal / nLairs, 1) 'average number of monsters per lair
            nMonsterSpawnChance(tabTempRS.Fields("Number")) = Round(nSpawnChance / nLairs, 2) 'average chance to spawn per lair
            'nTotalPossy = nTotalPossy + nMobsTotal 'nMonsterPossy(tabTempRS.Fields("Number"))
        End If
        
        tabTempRS.MoveNext
    Loop
    
    tabTempRS.MoveLast
    nAveragePossSpawns = Round(nTotalTOTALLairs / tabTempRS.RecordCount)
    nAverageMobsPerLair = Round(nTotalPossy / nTotalLairs, 1)
    
    tabTempRS.Close
    Set tabTempRS = Nothing
End If

'Debug.Print nTotalLairs

out:
On Error Resume Next
Set dictLairs = Nothing
Exit Sub
error:
Call HandleError("CalculateAverageLairs")
Resume out:
End Sub

Private Sub TestMonExpMulti()
On Error GoTo nomulti:
Dim sTest As String

'this just tests the exp multiplier field. if it exists, it wont error out

sTest = tabMonsters.Fields("ExpMulti").Value

UseExpMulti = True

Exit Sub
nomulti:
Err.clear
End Sub

Public Sub CloseDatabases()
On Error Resume Next

tabItems.Close
tabSpells.Close
tabRaces.Close
tabClasses.Close
tabInfo.Close
tabMonsters.Close
tabShops.Close
tabRooms.Close
tabTBInfo.Close

Set tabRooms = Nothing
Set tabMonsters = Nothing
Set tabShops = Nothing
Set tabItems = Nothing
Set tabSpells = Nothing
Set tabRaces = Nothing
Set tabClasses = Nothing
Set tabInfo = Nothing
Set tabTBInfo = Nothing

DB.Close
'WS.Close

Set DB = Nothing
'Set WS = Nothing

End Sub


Public Function GetShopName(ByVal nNum As Long, Optional ByVal bNoNumber As Boolean) As String
On Error GoTo error:

If nNum = 0 Then GetShopName = "None": Exit Function
If tabShops.RecordCount = 0 Then GetShopName = nNum: Exit Function

tabShops.Index = "pkShops"
tabShops.Seek "=", nNum
If Not tabShops.NoMatch Then
    GetShopName = tabShops.Fields("Name")
    If Not bNoNumber Then GetShopName = GetShopName & "(" & nNum & ")"
Else
    GetShopName = nNum
End If

Exit Function
error:
HandleError
GetShopName = nNum

End Function

Public Function GetItemShopRegenPCT(ByVal nShopNum As Long, ByVal nItemNum As Long) As Currency
Dim nRegenTimeMultiplier As Currency, x As Integer
On Error GoTo error:

GetItemShopRegenPCT = 0
If nItemNum < 1 Or nShopNum < 1 Then Exit Function

tabShops.Index = "pkShops"
tabShops.Seek "=", nShopNum
If tabShops.NoMatch = True Then Exit Function

If tabShops.Fields("ShopType") = 8 Then Exit Function
    
For x = 0 To 19
    If tabShops.Fields("Item-" & x) = nItemNum And tabShops.Fields("Max-" & x) > 0 Then
    
        If tabShops.Fields("Time-" & x) > 0 And tabShops.Fields("%-" & x) > 0 And tabShops.Fields("Amount-" & x) > 0 Then
            nRegenTimeMultiplier = 1440 / tabShops.Fields("Time-" & x)
            GetItemShopRegenPCT = GetItemShopRegenPCT + (tabShops.Fields("Amount-" & x) * nRegenTimeMultiplier * (tabShops.Fields("%-" & x) / 100))
        Else
            'stock only, we'll give it a 1% chance
            GetItemShopRegenPCT = GetItemShopRegenPCT + 1
        End If
    End If
Next

out:
On Error Resume Next
If GetItemShopRegenPCT > 99 Then GetItemShopRegenPCT = 99
GetItemShopRegenPCT = Round(GetItemShopRegenPCT, 2)
Exit Function
error:
Call HandleError("GetItemShopRegenPCT")
Resume out:
End Function

Public Function GetSpellName(ByVal nNum As Long, Optional ByVal bNoNumber As Boolean) As String
On Error GoTo error:

If nNum = 0 Then GetSpellName = "None": Exit Function
If tabSpells.RecordCount = 0 Then GetSpellName = nNum: Exit Function

tabSpells.Index = "pkSpells"
tabSpells.Seek "=", nNum
If tabSpells.NoMatch = True Then
    GetSpellName = nNum
Else
    GetSpellName = tabSpells.Fields("Name")
    If Not bNoNumber Then GetSpellName = GetSpellName & "(" & nNum & ")"
End If

Exit Function
error:
Call HandleError("GetSpellName")
GetSpellName = nNum

End Function

Public Function GetRaceHPBonus(ByVal nNum As Long) As Integer
On Error GoTo error:

If nNum = 0 Then GetRaceHPBonus = 0: Exit Function
If tabRaces.RecordCount = 0 Then GetRaceHPBonus = 0: Exit Function

tabRaces.Index = "pkRaces"
tabRaces.Seek "=", nNum
If tabRaces.NoMatch = True Then
    GetRaceHPBonus = 0
Else
    GetRaceHPBonus = tabRaces.Fields("HPPerLVL")
End If

Exit Function
error:
Call HandleError("GetRaceHPBonus")
GetRaceHPBonus = 0
End Function

Public Function GetClassMaxHP(ByVal nNum As Long) As Integer
On Error GoTo error:

If nNum = 0 Then GetClassMaxHP = 0: Exit Function
If tabClasses.RecordCount = 0 Then GetClassMaxHP = 0: Exit Function

tabClasses.Index = "pkClasses"
tabClasses.Seek "=", nNum
If tabClasses.NoMatch = True Then
    GetClassMaxHP = 0
Else
    GetClassMaxHP = tabClasses.Fields("MinHits") + tabClasses.Fields("MaxHits")
End If

Exit Function
error:
Call HandleError("GetClassMaxHP")
GetClassMaxHP = 0
End Function

Public Function GetClassMinHP(ByVal nNum As Long) As Integer
On Error GoTo error:

If nNum = 0 Then GetClassMinHP = 0: Exit Function
If tabClasses.RecordCount = 0 Then GetClassMinHP = 0: Exit Function

tabClasses.Index = "pkClasses"
tabClasses.Seek "=", nNum
If tabClasses.NoMatch = True Then
    GetClassMinHP = 0
Else
    GetClassMinHP = tabClasses.Fields("MinHits")
End If

Exit Function
error:
Call HandleError("GetClassMinHP")
GetClassMinHP = 0
End Function

Public Function GetClassName(ByVal nNum As Long) As String
On Error GoTo error:

If nNum = 0 Then GetClassName = "None": Exit Function
If tabClasses.RecordCount = 0 Then GetClassName = nNum: Exit Function

tabClasses.Index = "pkClasses"
tabClasses.Seek "=", nNum
If tabClasses.NoMatch = True Then
    GetClassName = nNum
Else
    GetClassName = tabClasses.Fields("Name")
End If

Exit Function
error:
Call HandleError("GetClassName")
GetClassName = nNum
End Function

Public Function GetClassMageryLVL(ByVal nNum As Long) As Integer

If nNum = 0 Then GetClassMageryLVL = 0: Exit Function
If tabClasses.RecordCount = 0 Then GetClassMageryLVL = 0: Exit Function

tabClasses.Index = "pkClasses"
tabClasses.Seek "=", nNum
If tabClasses.NoMatch = True Then
    GetClassMageryLVL = 0
Else
    GetClassMageryLVL = tabClasses.Fields("MageryLVL")
End If

Exit Function
error:
Call HandleError("GetClassMageryLVL")
GetClassMageryLVL = 0
End Function

Public Function GetClassMagery(ByVal nNum As Long) As enmMagicEnum

If nNum = 0 Then GetClassMagery = None: Exit Function
If tabClasses.RecordCount = 0 Then GetClassMagery = None: Exit Function

tabClasses.Index = "pkClasses"
tabClasses.Seek "=", nNum
If tabClasses.NoMatch = True Then
    GetClassMagery = None
Else
    Select Case tabClasses.Fields("MageryType")
        Case 1:
            GetClassMagery = Mage
        Case 2:
            GetClassMagery = Priest
        Case 3:
            GetClassMagery = Druid
        Case 4:
            GetClassMagery = Bard
        Case 5:
            GetClassMagery = Kai
        Case Else:
            GetClassMagery = None
    End Select
End If

Exit Function
error:
Call HandleError("GetClassMagery")
GetClassMagery = None
End Function

Public Function GetClassCombat(ByVal nNum As Long) As Integer
On Error GoTo error:

If nNum = 0 Then GetClassCombat = 1: Exit Function
If tabClasses.RecordCount = 0 Then GetClassCombat = 1: Exit Function

tabClasses.Index = "pkClasses"
tabClasses.Seek "=", nNum
If tabClasses.NoMatch = True Then
    GetClassCombat = 1
Else
    GetClassCombat = tabClasses.Fields("CombatLVL") - 2
End If

Exit Function
error:
Call HandleError("GetClassCombat")
GetClassCombat = 1
End Function

Public Function GetRaceName(ByVal nNum As Long) As String
On Error GoTo error:

If nNum = 0 Then GetRaceName = "None": Exit Function
If tabRaces.RecordCount = 0 Then GetRaceName = nNum: Exit Function

tabRaces.Index = "pkRaces"
tabRaces.Seek "=", nNum
If tabRaces.NoMatch = True Then
    GetRaceName = nNum
Else
    GetRaceName = tabRaces.Fields("Name")
End If

Exit Function
error:
Call HandleError("GetRaceName")
GetRaceName = nNum
End Function

Public Function GetRaceCP(ByVal nNum As Long) As Integer
On Error GoTo error:

If nNum = 0 Then GetRaceCP = 100: Exit Function
If tabRaces.RecordCount = 0 Then GetRaceCP = 100: Exit Function

tabRaces.Index = "pkRaces"
tabRaces.Seek "=", nNum
If tabRaces.NoMatch = True Then
    GetRaceCP = 100
Else
    GetRaceCP = tabRaces.Fields("BaseCP")
End If

Exit Function
error:
Call HandleError("GetRaceCP")
GetRaceCP = 100
End Function

Public Function GetRaceStealth(ByVal nNum As Long) As Boolean
Dim x As Integer
On Error GoTo error:

If nNum = 0 Then Exit Function
If tabRaces.RecordCount = 0 Then Exit Function

tabRaces.Index = "pkRaces"
tabRaces.Seek "=", nNum
If tabRaces.NoMatch = True Then Exit Function

For x = 0 To 9
    If tabRaces.Fields("Abil-" & x) = 102 Then
        GetRaceStealth = True
        Exit For
    End If
Next x

Exit Function
error:
Call HandleError("GetRaceStealth")
End Function

Public Function GetClassStealth(ByVal nNum As Long) As Boolean
Dim x As Integer
On Error GoTo error:

If nNum = 0 Then Exit Function
If tabClasses.RecordCount = 0 Then Exit Function

tabClasses.Index = "pkClasses"
tabClasses.Seek "=", nNum
If tabClasses.NoMatch = True Then Exit Function

For x = 0 To 9
    If tabClasses.Fields("Abil-" & x) = 103 Then
        GetClassStealth = True
        Exit For
    End If
Next x

Exit Function
error:
Call HandleError("GetClassStealth")
End Function

Public Function GetMultiMonsterNames(ByVal sNumbers As String, ByVal HideNumber As Boolean) As String
Dim x As Long, y As Long
On Error GoTo error:

If sNumbers = "" Then GetMultiMonsterNames = "None": Exit Function
If tabMonsters.RecordCount = 0 Then Exit Function

tabMonsters.Index = "pkMonsters"
x = 0
Do While Not InStr(x + 1, sNumbers, ",") = 0
    y = InStr(x + 1, sNumbers, ",")
    
    tabMonsters.Seek "=", Val(Mid(sNumbers, x + 1, y - x - 1))
    If tabItems.NoMatch = False Then
        GetMultiMonsterNames = GetMultiMonsterNames & IIf(GetMultiMonsterNames = "", "", ", ") _
            & tabMonsters.Fields("Name")
            
        If Not HideNumber Then
            GetMultiMonsterNames = GetMultiMonsterNames & "(" & tabMonsters.Fields("Number") & ")"
        End If
    End If
    x = y
Loop

Exit Function
error:
Call HandleError("GetMultiMonsterNames")
GetMultiMonsterNames = sNumbers
End Function
Public Function GetMonsterName(ByVal nNum As Long, ByVal bNoNumber As Boolean) As String
On Error GoTo error:
GetMonsterName = nNum

If nNum = 0 Then GetMonsterName = "None": Exit Function
If tabMonsters.RecordCount = 0 Then Exit Function

tabMonsters.Index = "pkMonsters"
tabMonsters.Seek "=", nNum
If tabMonsters.NoMatch = True Then
    GetMonsterName = nNum
Else
    GetMonsterName = tabMonsters.Fields("Name")
    If Not bNoNumber Then GetMonsterName = GetMonsterName & "(" & nNum & ")"
End If


Exit Function
error:
Call HandleError("GetMonsterName")
End Function
Public Function GetRoomName(Optional ByVal sMapRoom As String, Optional ByVal nMap As Long, _
    Optional ByVal nRoom As Long, Optional bNoRoomNumber As Boolean) As String
Dim tExit As RoomExitType, sName As String

If sMapRoom = "" Then
    tExit.Map = nMap
    tExit.Room = nRoom
Else
    tExit = ExtractMapRoom(sMapRoom)
End If

If tExit.Map = 0 Or tExit.Room = 0 Then GetRoomName = "?": Exit Function

tabRooms.Index = "idxRooms"
tabRooms.Seek "=", tExit.Map, tExit.Room
If tabRooms.NoMatch = True Then
    GetRoomName = tExit.Map & "/" & tExit.Room
Else
    sName = tabRooms.Fields("Name")
    If sName = "" Then sName = "(no name)"
    If Not bNoRoomNumber Then sName = sName & " (" & tExit.Map & "/" & tExit.Room & ")"
    GetRoomName = sName
End If

End Function

Public Function GetRoomCMDTB(Optional ByVal sMapRoom As String, Optional ByVal nMap As Long, Optional ByVal nRoom As Long) As Long
Dim tExit As RoomExitType

If sMapRoom = "" Then
    tExit.Map = nMap
    tExit.Room = nRoom
Else
    tExit = ExtractMapRoom(sMapRoom)
End If

If tExit.Map = 0 Or tExit.Room = 0 Then GetRoomCMDTB = 0: Exit Function

tabRooms.Index = "idxRooms"
tabRooms.Seek "=", tExit.Map, tExit.Room
If tabRooms.NoMatch = True Then
    GetRoomCMDTB = 0
Else
    GetRoomCMDTB = tabRooms.Fields("CMD")
End If

End Function

Public Function GetItemLimit(ByVal nItemNumber As Long) As Integer
On Error GoTo error:

GetItemLimit = -1

If nItemNumber = 0 Then Exit Function
    
If Not tabItems.Fields("Number") = nItemNumber Then
    tabItems.Index = "pkItems"
    tabItems.Seek "=", nItemNumber
    If tabItems.NoMatch = True Then
        tabItems.MoveFirst
        Exit Function
    End If
End If

GetItemLimit = tabItems.Fields("Limit")

out:
Exit Function
error:
Call HandleError("GetItem")
GetItemLimit = -1
Resume out:
End Function

Public Function ItemHasAbility(ByVal nItemNumber As Long, ByVal nAbility As Integer) As Integer
Dim x As Integer
On Error GoTo error:

'-1 = does not have
'>=0 = value of ability

ItemHasAbility = -1
If nAbility <= 0 Or nItemNumber <= 0 Then Exit Function

If Not tabItems.Fields("Number") = nItemNumber Then
    tabItems.Index = "pkItems"
    tabItems.Seek "=", nItemNumber
    If tabItems.NoMatch Then
        tabItems.MoveFirst
        Exit Function
    End If
End If

For x = 0 To 9
    If tabItems.Fields("Abil-" & x) = nAbility Then
        ItemHasAbility = tabItems.Fields("AbilVal-" & x)
        Exit Function
    End If
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("ItemHasAbility")
Resume out:
End Function

Public Function ItemIsChest(ByVal nItemNumber As Long) As Boolean
On Error GoTo error:

If Not tabItems.Fields("Number") = nItemNumber Then
    tabItems.Index = "pkItems"
    tabItems.Seek "=", nItemNumber
    If tabItems.NoMatch Then
        tabItems.MoveFirst
        Exit Function
    End If
End If

If tabItems.Fields("ItemType") = 8 Then ItemIsChest = True
out:
On Error Resume Next
Exit Function
error:
Call HandleError("ItemIsChest")
Resume out:
End Function

Public Function GetItemCost(ByVal nNum As Long, Optional ByVal MarkUp As Integer) As typItemCostDetail
On Error GoTo error:

If nNum = 0 Or tabItems.RecordCount = 0 Then
    GetItemCost.Cost = 0
    GetItemCost.Coin = 0
    Exit Function
End If

If Not tabItems.Fields("Number") = nNum Then
    tabItems.Index = "pkItems"
    tabItems.Seek "=", nNum
    If tabItems.NoMatch = True Then
        GetItemCost.Cost = 0
        GetItemCost.Coin = 0
        tabItems.MoveFirst
        Exit Function
    End If
End If

If tabItems.Fields("Price") = 0 Then
    GetItemCost.Cost = 0
    GetItemCost.Coin = 0
Else
    If MarkUp > 0 Then
        GetItemCost.Cost = tabItems.Fields("Price") + Fix(tabItems.Fields("Price") * (MarkUp / 100))
    Else
        GetItemCost.Cost = tabItems.Fields("Price")
    End If
    
    GetItemCost.Coin = tabItems.Fields("Currency")
End If


Exit Function
error:
HandleError
GetItemCost.Cost = 0
GetItemCost.Coin = 0
End Function

Public Function GetItemWeight(ByVal nNum As Long) As Long
On Error GoTo error:

If nNum = 0 Then Exit Function

tabItems.Index = "pkItems"
tabItems.Seek "=", nNum
If tabItems.NoMatch = True Then
    tabItems.MoveFirst
    Exit Function
End If
GetItemWeight = tabItems.Fields("Encum")

Exit Function
error:
Call HandleError("GetItemWeight")
End Function

'Public Function GetItemCostType(ByVal nNum As Long) As Integer
'
'On Error GoTo Error:
'
'If nNum = 0 Then GetItemCostType = 0: Exit Function
'If tabItems.RecordCount = 0 Then GetItemCostType = 0: Exit Function
'
'If Not tabItems.Fields("Number") = nNum Then
'    tabItems.Index = "pkItems"
'    tabItems.Seek "=", nNum
'    If tabItems.NoMatch = True Then
'        GetItemCostType = 0
'        Exit Function
'    End If
'End If
'
'GetItemCostType = tabItems.Fields("Currency")
'
'Exit Function
'Error:
'Call HandleError("GetItemCostType")
'
'End Function

Public Function GetItemName(ByVal nNum As Long, Optional ByVal bNoNumber As Boolean) As String
On Error GoTo error:

If nNum = 0 Then GetItemName = "None": Exit Function
If tabItems.RecordCount = 0 Then GetItemName = nNum: Exit Function

tabItems.Index = "pkItems"
tabItems.Seek "=", nNum
If tabItems.NoMatch = True Then
    tabItems.MoveFirst
    GetItemName = nNum
Else
    GetItemName = tabItems.Fields("Name")
    If Not bNoNumber Then GetItemName = GetItemName & "(" & nNum & ")"
End If

Exit Function
error:
HandleError
GetItemName = nNum
End Function

Public Function GetCurrentSpellMinMax(Optional ByRef bUseLevel As Boolean, Optional ByVal nLevel As Integer, Optional ByRef bNoHeader As Boolean) As SpellMinMaxDur
Dim nMin As Currency, nMinIncr As Currency, nMinLVLs As Currency
Dim nMax As Currency, nMaxIncr As Currency, nMaxLVLs As Currency
Dim nDur As Currency, nDurIncr As Currency, nDurLVLs As Currency
Dim sMin As String, sMax As String, sDur As String
On Error GoTo error:

If tabSpells Is Nothing Then Exit Function
If tabSpells.EOF Then Exit Function

nMin = tabSpells.Fields("MinBase")
nMinIncr = tabSpells.Fields("MinInc")
nMinLVLs = tabSpells.Fields("MinIncLVLs")

nMax = tabSpells.Fields("MaxBase")
nMaxIncr = tabSpells.Fields("MaxInc")
nMaxLVLs = tabSpells.Fields("MaxIncLVLs")

nDur = tabSpells.Fields("Dur")
nDurIncr = tabSpells.Fields("DurInc")
nDurLVLs = tabSpells.Fields("DurIncLVLs")

If bUseLevel Then
    If (nMinIncr = 0 Or nMinLVLs = 0) And (nMaxIncr = 0 Or nMaxLVLs = 0) And _
        (nDurIncr = 0 Or nDurLVLs = 0) Then bUseLevel = False
End If

If tabSpells.Fields("Cap") = 0 And tabSpells.Fields("ReqLevel") = 0 And bUseLevel = False Then
    sDur = nDur
    sMax = nMax
    sMin = nMin
Else
    'if there is an amount specified in the ability value, dont use the spells min and max
    'If Not tabSpells.Fields("Ability Value 0") = 0 Then
    '    sMin = tabSpells.Fields("Ability Value 0")
    '    sMax = tabSpells.Fields("Ability Value 0")
    '    GoTo CalcDur:
    'End If
    
    'figure out mins and maxs...
    If nMinLVLs = 0 Or nMinIncr = 0 Then
        sMin = nMin
    Else
        If bUseLevel = True Then
            nMin = nMin + (Round(nMinIncr / nMinLVLs, 2) * nLevel)
            nMin = Fix(nMin)
            sMin = nMin
        Else
            bNoHeader = True
            sMin = nMin & "+(" & Round(nMinIncr / nMinLVLs, 2) & "*lvl)"
        End If
    End If
    
    If nMaxLVLs = 0 Or nMaxIncr = 0 Then
        sMax = nMax
    Else
        If bUseLevel = True Then
            nMax = nMax + (Round(nMaxIncr / nMaxLVLs, 2) * nLevel)
            nMax = Fix(nMax)
            sMax = nMax
        Else
            bNoHeader = True
            sMax = nMax & "+(" & Round(nMaxIncr / nMaxLVLs, 2) & "*lvl)"
        End If
    End If
    
'CalcDur:
    If nDurLVLs = 0 Or nDurIncr = 0 Then
        sDur = nDur
    Else
        If bUseLevel = True Then
            nDur = nDur + (Round(nDurIncr / nDurLVLs, 2) * nLevel)
            nDur = Fix(nDur)
            sDur = nDur
        Else
            sDur = nDur & "+(" & Round(nDurIncr / nDurLVLs, 2) & "*lvl)"
        End If
    End If
End If

out:
On Error Resume Next
GetCurrentSpellMinMax.nMin = nMin
GetCurrentSpellMinMax.nMax = nMax
GetCurrentSpellMinMax.nDur = nDur
GetCurrentSpellMinMax.sMin = sMin
GetCurrentSpellMinMax.sMax = sMax
GetCurrentSpellMinMax.sDur = sDur
Exit Function
error:
Call HandleError("GetSpellMinMax")
Resume out:
End Function
Public Function PullSpellEQ(ByVal bCalcLevel As Boolean, Optional ByVal nLevel As Integer, _
    Optional ByVal nSpell As Long, Optional ByRef LV As ListView, Optional bMinMaxDamageOnly As Boolean = False, _
    Optional bForMonster As Boolean, Optional ByVal bPercentColumn As Boolean) As String
Dim oLI As ListItem, sTemp As String
Dim sMin As String, sMax As String, sDur As String, sExtra As String
Dim nMin As Currency, nMax As Currency, nDur As Currency, tSpellMinMaxDur As SpellMinMaxDur
Dim sMinHeader As String, sMaxHeader As String, sRemoves As String, bUseLevel As Boolean
Dim y As Long, nAbilValue As Long, x As Integer, bNoHeader As Boolean, nMap As Long
Dim bDoesDamage As Boolean

On Error GoTo error:

nSpellNest = nSpellNest + 1

If tabSpells.RecordCount = 0 Then PullSpellEQ = "(No Spell Records)": GoTo out:
If nSpellNest > 19 Then PullSpellEQ = " ... to infinity and beyond?": GoTo out:

If bQuickSpell And nSpellNest > 1 Then
    PullSpellEQ = "(click)"
    GoTo out:
End If

'base + ((how_much_incr / lvls_for_incr) * level)

If Not nSpell = 0 Then
    tabSpells.Index = "pkSpells"
    tabSpells.Seek "=", nSpell
    If tabSpells.NoMatch = True Then
        PullSpellEQ = "?"
        GoTo out:
    End If
Else
    nSpell = tabSpells.Fields("Number")
End If

bUseLevel = bCalcLevel
If bUseLevel Then
    'use the value in the global filter for level if a level hasn't been specified
    If nLevel = 0 And frmMain.chkGlobalFilter.Value = 1 Then
        nLevel = Val(frmMain.txtGlobalLevel(0).Text)
    End If
    
    'make the level less if it's above the level cap, and more if it's below the required, except for monster attacks
    If Not bForMonster Then
        If nLevel > tabSpells.Fields("Cap") And tabSpells.Fields("Cap") > 0 Then nLevel = tabSpells.Fields("Cap")
        If nLevel < tabSpells.Fields("ReqLevel") Then nLevel = tabSpells.Fields("ReqLevel")
    End If
    If nLevel < 1 Then nLevel = tabSpells.Fields("ReqLevel")
    
    If nLevel = 0 Then bUseLevel = False
End If

tSpellMinMaxDur = GetCurrentSpellMinMax(bUseLevel, nLevel, bNoHeader)

nMin = tSpellMinMaxDur.nMin
nMax = tSpellMinMaxDur.nMax
nDur = tSpellMinMaxDur.nDur
sMin = tSpellMinMaxDur.sMin
sMax = tSpellMinMaxDur.sMax
sDur = tSpellMinMaxDur.sDur

For x = 0 To 9
    If Not tabSpells.Fields("Abil-" & x) = 0 Then
    
        Select Case tabSpells.Fields("Abil-" & x)
            Case 1, 8, 17, 18, 19:
                bDoesDamage = True
                If bMinMaxDamageOnly Then Exit For
        End Select
        
        sMinHeader = ""
        sMaxHeader = ""
        nAbilValue = tabSpells.Fields("AbilVal-" & x)
        'If nAbilValue = 0 And nMin = nMax Then nAbilValue = nMin
        
        If tabSpells.Fields("Abil-" & x) = 122 Then 'RemoveSpell
            If bQuickSpell Then
                If sRemoves = "" Then sRemoves = "click"
            Else
                If Not sRemoves = "" Then sRemoves = sRemoves & ", "
                sRemoves = sRemoves & GetSpellName(nAbilValue, bHideRecordNumbers)
            End If
            
        Else
            If Not sExtra = "" Then sExtra = sExtra & ", "
            
            If nAbilValue = 0 Then
                Select Case tabSpells.Fields("Abil-" & x)
                    Case 140: 'teleport
                        sExtra = sExtra & GetAbilityStats(tabSpells.Fields("Abil-" & x), , IIf(LV Is Nothing, Nothing, LV), , bPercentColumn) _
                            & " " & IIf(sMin = sMax, sMin, sMin & " to " & sMax)
                        If Not LV Is Nothing Then
                            nMap = 0
                            For y = 0 To 9
                                If tabSpells.Fields("Abil-" & y) = 141 Then 'tele map
                                    nMap = tabSpells.Fields("AbilVal-" & y)
                                End If
                            Next y
                            
                            If nMap > 0 Then
                                For y = Val(sMin) To Val(sMax)
                                    If bPercentColumn Then
                                        Set oLI = LV.ListItems.Add()
                                        oLI.Text = ""
                                        oLI.ListSubItems.Add , , "Teleport: " & GetRoomName(, nMap, y, False)
                                        oLI.ListSubItems(1).Tag = nMap & "/" & y
                                    Else
                                        Set oLI = LV.ListItems.Add(, , "Teleport: " & GetRoomName(, nMap, y, False))
                                        oLI.Tag = nMap & "/" & y
                                    End If
                                    Set oLI = Nothing
                                Next y
                            End If
                        End If
                    Case 148: 'textblock
                        sExtra = sExtra & GetAbilityStats(tabSpells.Fields("Abil-" & x), , LV, , bPercentColumn) _
                            & " " & IIf(sMin = sMax, sMin, sMin & " to " & sMax)
                        If Not LV Is Nothing Then
                            For y = Val(sMin) To Val(sMax)
                                If bPercentColumn Then
                                    Set oLI = LV.ListItems.Add()
                                    oLI.Text = ""
                                    oLI.ListSubItems.Add , , "Execute: Textblock " & y
                                    oLI.ListSubItems(1).Tag = y
                                Else
                                    Set oLI = LV.ListItems.Add(, , "Execute: Textblock " & y)
                                    oLI.Tag = y
                                End If
                                Set oLI = Nothing
                            Next y
                        End If
                    Case 151: 'endcast
                        If bQuickSpell Then
                            If nMax > nMin Then
                                sExtra = sExtra & "End cast " & nMin & " to " & nMax
                            Else
                                sExtra = sExtra & "End cast " & nMin
                            End If
                        Else
                            If nMin >= nMax Then
                                sExtra = sExtra & "EndCast [" & GetSpellName(nMin, bHideRecordNumbers) & ", " & PullSpellEQ(bCalcLevel, nLevel, nMin, LV, , , bPercentColumn) & "]"
                            Else
                                sExtra = sExtra & "EndCast [{" & GetSpellName(nMin, bHideRecordNumbers) & ", " & PullSpellEQ(bCalcLevel, nLevel, nMin, LV, , , bPercentColumn) & "}"
                                For y = nMin + 1 To nMax
                                    sExtra = sExtra & " OR {" & GetSpellName(y, bHideRecordNumbers) & ", " & PullSpellEQ(bCalcLevel, nLevel, y, LV, , , bPercentColumn) & "}"
                                Next y
                                sExtra = sExtra & "]"
                            End If
                        End If
                        
'                    Case 124: 'negateabil
'                        If sMin >= sMax Then
'                            sExtra = sExtra & "NegateAbility " & GetAbilityName(sMin)
'                        Else
'                            sExtra = sExtra & "NegateAbilities{" & GetAbilityName(sMin)
'                            For y = sMin + 1 To sMax
'                                sExtra = sExtra & " OR " & GetAbilityName(y)
'                            Next y
'                            sExtra = sExtra & "}"
'                        End If
                    Case 12: 'summon
                        If bQuickSpell Then
                            sExtra = sExtra & "Summon"
                        Else
                            If nMin >= nMax Then
                                sTemp = GetMonsterName(nMin, bHideRecordNumbers)
                                sExtra = sExtra & "Summon " & sTemp
                                If Not LV Is Nothing Then
                                    Set oLI = LV.ListItems.Add()
                                    If bPercentColumn Then
                                        oLI.Text = ""
                                        oLI.ListSubItems.Add , , "Summon: " & sTemp
                                        oLI.ListSubItems(1).Tag = nMin
                                    Else
                                        oLI.Text = "Summon: " & sTemp
                                        oLI.Tag = nMin
                                    End If
                                    Set oLI = Nothing
                                End If
                            Else
                                sTemp = GetMonsterName(nMin, bHideRecordNumbers)
                                sExtra = sExtra & "Summons{" & sTemp
                                If Not LV Is Nothing Then
                                    Set oLI = LV.ListItems.Add()
                                    If bPercentColumn Then
                                        oLI.Text = ""
                                        oLI.ListSubItems.Add , , "Summon: " & sTemp
                                        oLI.ListSubItems(1).Tag = nMin
                                    Else
                                        oLI.Text = "Summon: " & sTemp
                                        oLI.Tag = nMin
                                    End If
                                    Set oLI = Nothing
                                End If
                                
                                For y = nMin + 1 To nMax
                                    sTemp = GetMonsterName(y, bHideRecordNumbers)
                                    sExtra = sExtra & " OR " & sTemp
                                    If Not LV Is Nothing Then
                                        Set oLI = LV.ListItems.Add()
                                        If bPercentColumn Then
                                            oLI.Text = ""
                                            oLI.ListSubItems.Add , , "Summon: " & sTemp
                                            oLI.ListSubItems(1).Tag = y
                                        Else
                                            oLI.Text = "Summon: " & sTemp
                                            oLI.Tag = y
                                        End If
                                        Set oLI = Nothing
                                    End If
                                Next y
                                sExtra = sExtra & "}"
                            End If
                        End If
                    Case 23, 51, 52, 80, 97, 98, 100, 108 To 113, 119, 138, 144:
                        '23 - effectsundead
                        '51: 'anti magic
                        '52: 'evil in combat
                        '80: 'effects animal
                        '97-98 - good/evil only
                        '100: 'loyal
                        '108: 'effects living
                        '109 To 113: 'nonliving, notgood, notevil, neutral, not neutral
                        '112 - neut only
                        '119: 'del@main
                        '138: 'roomvis
                        '144: 'non magic spell
                        sExtra = sExtra & GetAbilityStats(tabSpells.Fields("Abil-" & x))
                    Case 7: 'DR
                        If Not bNoHeader Then
                            If Val(sMin) > 0 Then sMinHeader = "+"
                            If Val(sMax) > 0 Then sMaxHeader = "+"
                        End If
                        
                        If bUseLevel Then
                            sExtra = sExtra & GetAbilityStats(tabSpells.Fields("Abil-" & x), , LV, , bPercentColumn) _
                                & " " & IIf(nMin = nMax, sMinHeader & (nMin / 10), sMinHeader & (nMin / 10) & " to " & sMaxHeader & (nMax / 10))
                        Else
                            sExtra = sExtra & GetAbilityStats(tabSpells.Fields("Abil-" & x), , LV, , bPercentColumn) _
                                & " " & IIf(sMin = sMax, sMinHeader & sMin, sMinHeader & sMin & " to " & sMaxHeader & sMax)
                        End If
                    Case Else:

                        If Not bNoHeader Then
                            Select Case tabSpells.Fields("Abil-" & x)
                                Case 1, 8, 17, 18, 19, 140, 141, 148:
                                'damage, drain, damage(on armr), poison, heal, teleport room, teleport map, textblocks
                                ' *** ALSO ADD THESE TO GetAbilityStats ***
                                Case Else:
                                    If Val(sMin) > 0 Then sMinHeader = "+"
                                    If Val(sMax) > 0 Then sMaxHeader = "+"
                            End Select
                        End If
                        
                        'sExtra = sExtra & GetAbilityStats(tabSpells.Fields("Abil-" & x), , IIf(LV Is Nothing, Nothing, LV)) _
                            & " " & IIf(sMin = sMax, sMinHeader & sMin, sMinHeader & sMin & " to " & sMaxHeader & sMax)
                        sExtra = sExtra & GetAbilityStats(tabSpells.Fields("Abil-" & x), , LV, bCalcLevel, bPercentColumn) _
                            & " " & IIf(sMin = sMax, sMinHeader & sMin, sMinHeader & sMin & " to " & sMaxHeader & sMax)
                        
                End Select
            Else
                Select Case tabSpells.Fields("Abil-" & x)
                    Case 148: 'textblock
                        sExtra = sExtra & GetAbilityStats(tabSpells.Fields("Abil-" & x), nAbilValue, IIf(LV Is Nothing, Nothing, LV), , bPercentColumn)
                        If Not LV Is Nothing Then
                            If bPercentColumn Then
                                Set oLI = LV.ListItems.Add()
                                oLI.Text = ""
                                oLI.ListSubItems.Add , , "Execute: Textblock " & nAbilValue
                                oLI.ListSubItems(1).Tag = nAbilValue
                            Else
                                Set oLI = LV.ListItems.Add(, , "Execute: Textblock " & nAbilValue)
                                oLI.Tag = nAbilValue
                            End If
                            Set oLI = Nothing
                        End If
                    Case 12: 'summon
                        If bQuickSpell Then
                            sExtra = sExtra & "Summon"
                        Else
                            sTemp = GetMonsterName(nAbilValue, bHideRecordNumbers)
                            sExtra = sExtra & "Summon " & sTemp
                            If Not LV Is Nothing Then
                                Set oLI = LV.ListItems.Add()
                                If bPercentColumn Then
                                    oLI.Text = ""
                                    oLI.ListSubItems.Add , , "Summon: " & sTemp
                                    oLI.ListSubItems(1).Tag = nAbilValue
                                Else
                                    oLI.Text = "Summon: " & sTemp
                                    oLI.Tag = nAbilValue
                                End If
                                Set oLI = Nothing
                            End If
                        End If
                        
                    Case 140: 'teleport
                        sExtra = sExtra & GetAbilityStats(tabSpells.Fields("Abil-" & x), nAbilValue, IIf(LV Is Nothing, Nothing, LV), , bPercentColumn)
                        If Not LV Is Nothing Then
                            nMap = 0
                            For y = 0 To 9
                                If tabSpells.Fields("Abil-" & y) = 141 Then
                                    nMap = tabSpells.Fields("AbilVal-" & y)
                                End If
                            Next y
                            
                            If nMap > 0 Then
                                If bPercentColumn Then
                                    Set oLI = LV.ListItems.Add()
                                    oLI.Text = ""
                                    oLI.ListSubItems.Add , , "Teleport: " & GetRoomName(, nMap, nAbilValue, False)
                                    oLI.ListSubItems(1).Tag = nMap & "/" & nAbilValue
                                Else
                                    Set oLI = LV.ListItems.Add(, , "Teleport: " & GetRoomName(, nMap, nAbilValue, False))
                                    oLI.Tag = nMap & "/" & nAbilValue
                                End If
                                Set oLI = Nothing
                            End If
                        End If
                    Case Else:
                        sExtra = sExtra & GetAbilityStats(tabSpells.Fields("Abil-" & x), nAbilValue, LV, bCalcLevel, bPercentColumn)
                        
                End Select
            End If
            
            If Right(sExtra, 2) = ", " Then sExtra = Left(sExtra, Len(sExtra) - 2)
        End If
        
        'reposition in case the ability function changed it
        If Not tabSpells.Fields("Number") = nSpell Then tabSpells.Seek "=", nSpell
    End If
Next x

If bMinMaxDamageOnly Then
    If bDoesDamage Then
        PullSpellEQ = sMin & ":" & sMax & IIf(nDur > 0, ":" & sDur, "")
    Else
        PullSpellEQ = "0:0:0"
    End If
    GoTo out:
End If

If sExtra = "" And sRemoves = "" Then
    PullSpellEQ = "(No EQ)"
    GoTo out:
End If

PullSpellEQ = sExtra

If bUseLevel = True Then
    If tabSpells.Fields("Cap") > 0 Or tabSpells.Fields("ReqLevel") > 0 Then
        PullSpellEQ = "(@lvl " & nLevel & "): " & PullSpellEQ
    End If
End If

'If Not sExtra = "" Then
'    PullSpellEQ = PullSpellEQ & ", " & sExtra
'End If

If Not sDur = "0" Then
    If Not PullSpellEQ = "" Then PullSpellEQ = PullSpellEQ & ", "
    PullSpellEQ = PullSpellEQ & "for " & sDur & " rounds"
End If

If Not sRemoves = "" Then
    If Not PullSpellEQ = "" Then PullSpellEQ = PullSpellEQ & " -- "
    PullSpellEQ = PullSpellEQ & "RemovesSpells(" & sRemoves & ")"
End If

out:
On Error Resume Next
nSpellNest = nSpellNest - 1
Exit Function

error:
Call HandleError("PullSpellEQ")
Resume out:
End Function

Public Function GetSpellMinDamage(ByVal nSpellNumber As Long, Optional ByVal nCastLevel As Integer, Optional nEnergyRem As Integer, Optional bForMonster As Boolean) As Long
Dim bDoesDamage As Boolean, x As Integer, nEndCast As Long
On Error GoTo error:

GetSpellMinDamage = 0
If nSpellNumber = 0 Then Exit Function

If Not tabSpells.Fields("Number") = nSpellNumber Then
    tabSpells.Index = "pkSpells"
    tabSpells.Seek "=", nSpellNumber
    If tabSpells.NoMatch = True Then Exit Function
End If

For x = 0 To 9
    Select Case tabSpells.Fields("Abil-" & x)
        Case 1, 8, 17: 'dmg/drain/dmg-mr
            bDoesDamage = True
            If tabSpells.Fields("AbilVal-" & x) > 0 Then
                GetSpellMinDamage = tabSpells.Fields("AbilVal-" & x)
            End If
        Case 151:
            nEndCast = tabSpells.Fields("AbilVal-" & x)
    End Select
Next x
If GetSpellMinDamage > 0 Then GoTo multi_calc:
If bDoesDamage = False Then Exit Function

If Not bForMonster Or nCastLevel = 0 Then
    If nCastLevel > tabSpells.Fields("Cap") And tabSpells.Fields("Cap") > 0 Then nCastLevel = tabSpells.Fields("Cap")
    If nCastLevel < tabSpells.Fields("ReqLevel") Then nCastLevel = tabSpells.Fields("ReqLevel")
End If

If tabSpells.Fields("MinIncLVLs") = 0 Or nCastLevel < 1 Then
    GetSpellMinDamage = tabSpells.Fields("MinBase")
Else
    GetSpellMinDamage = tabSpells.Fields("MinBase") + Fix((tabSpells.Fields("MinInc") / tabSpells.Fields("MinIncLVLs")) * nCastLevel)
End If

multi_calc:
If bForMonster Then Exit Function

If nEnergyRem = 0 Then nEnergyRem = 1000
nEnergyRem = nEnergyRem - tabSpells.Fields("EnergyCost")
If nEnergyRem < 1 Then nEnergyRem = 1

If nEnergyRem >= 143 And tabSpells.Fields("EnergyCost") >= 143 Then
    If nEndCast = 0 Then
        If tabSpells.Fields("EnergyCost") <= 500 Then
            GetSpellMinDamage = GetSpellMinDamage + (GetSpellMinDamage * Fix(nEnergyRem / tabSpells.Fields("EnergyCost")))
        End If
    Else
        GetSpellMinDamage = GetSpellMinDamage + GetSpellMinDamage(nEndCast, nCastLevel, nEnergyRem, bForMonster)
    End If
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetSpellMinDamage")
Resume out:
End Function

Public Function GetSpellMaxDamage(ByVal nSpellNumber As Long, Optional ByVal nCastLevel As Integer, Optional nEnergyRem As Integer, Optional bForMonster As Boolean) As Long
Dim bDoesDamage As Boolean, x As Integer, nEndCast As Long
On Error GoTo error:

GetSpellMaxDamage = 0
If nSpellNumber = 0 Then Exit Function

If Not tabSpells.Fields("Number") = nSpellNumber Then
    tabSpells.Index = "pkSpells"
    tabSpells.Seek "=", nSpellNumber
    If tabSpells.NoMatch = True Then Exit Function
End If

For x = 0 To 9
    Select Case tabSpells.Fields("Abil-" & x)
        Case 1, 8, 17: 'dmg/drain/dmg-mr
            bDoesDamage = True
            If tabSpells.Fields("AbilVal-" & x) > 0 Then
                GetSpellMaxDamage = tabSpells.Fields("AbilVal-" & x)
            End If
        Case 151:
            nEndCast = tabSpells.Fields("AbilVal-" & x)
    End Select
Next x
If GetSpellMaxDamage > 0 Then GoTo multi_calc:
If bDoesDamage = False Then Exit Function

If Not bForMonster Or nCastLevel = 0 Then
    If nCastLevel > tabSpells.Fields("Cap") And tabSpells.Fields("Cap") > 0 Then nCastLevel = tabSpells.Fields("Cap")
    If nCastLevel < tabSpells.Fields("ReqLevel") Then nCastLevel = tabSpells.Fields("ReqLevel")
End If

If tabSpells.Fields("MaxIncLVLs") = 0 Or nCastLevel < 1 Then
    GetSpellMaxDamage = tabSpells.Fields("MaxBase")
Else
    GetSpellMaxDamage = tabSpells.Fields("MaxBase") + Fix((tabSpells.Fields("MaxInc") / tabSpells.Fields("MaxIncLVLs")) * nCastLevel)
End If

multi_calc:
If bForMonster Then Exit Function

If nEnergyRem = 0 Then nEnergyRem = 1000
nEnergyRem = nEnergyRem - tabSpells.Fields("EnergyCost")
If nEnergyRem < 1 Then nEnergyRem = 1

If nEnergyRem >= 143 And tabSpells.Fields("EnergyCost") >= 143 Then
    If nEndCast = 0 Then
        If tabSpells.Fields("EnergyCost") <= 500 Then
            GetSpellMaxDamage = GetSpellMaxDamage + (GetSpellMaxDamage * Fix(nEnergyRem / tabSpells.Fields("EnergyCost")))
        End If
    Else
        GetSpellMaxDamage = GetSpellMaxDamage + GetSpellMaxDamage(nEndCast, nCastLevel, nEnergyRem, bForMonster)
    End If
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetSpellMaxDamage")
Resume out:
End Function

Public Function GetSpellDuration(ByVal nSpellNumber As Long, Optional ByVal nCastLevel As Integer, Optional bForMonster As Boolean) As Long
On Error GoTo error:

GetSpellDuration = 0
If nSpellNumber = 0 Then Exit Function

If Not tabSpells.Fields("Number") = nSpellNumber Then
    tabSpells.Index = "pkSpells"
    tabSpells.Seek "=", nSpellNumber
    If tabSpells.NoMatch = True Then Exit Function
End If

If Not bForMonster Or nCastLevel = 0 Then
    If nCastLevel > tabSpells.Fields("Cap") And tabSpells.Fields("Cap") > 0 Then nCastLevel = tabSpells.Fields("Cap")
    If nCastLevel < tabSpells.Fields("ReqLevel") Then nCastLevel = tabSpells.Fields("ReqLevel")
End If

If tabSpells.Fields("DurIncLVLs") = 0 Or nCastLevel < 1 Then
    GetSpellDuration = tabSpells.Fields("Dur")
Else
    GetSpellDuration = tabSpells.Fields("Dur") + Fix((tabSpells.Fields("DurInc") / tabSpells.Fields("DurIncLVLs")) * nCastLevel)
End If

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetSpellDuration")
Resume out:
End Function

Public Function SpellIsAreaAttack(ByVal nSpellNumber As Long) As Boolean
On Error GoTo error:

SpellIsAreaAttack = False

If Not tabSpells.Fields("Number") = nSpellNumber Then
    tabSpells.Index = "pkSpells"
    tabSpells.Seek "=", nSpellNumber
    If tabSpells.NoMatch Then
        tabSpells.MoveFirst
        Exit Function
    End If
End If

Select Case tabSpells.Fields("Targets")
    Case 9, 11, 12:
        SpellIsAreaAttack = True
End Select

out:
On Error Resume Next
Exit Function
error:
Call HandleError("SpellIsAreaAttack")
Resume out:
End Function

Public Function SpellHasAbility(ByVal nSpellNumber As Long, ByVal nAbility As Integer) As Integer
Dim x As Integer
On Error GoTo error:

'-1 = does not have
'>=0 = value of ability

SpellHasAbility = -1
If nAbility <= 0 Or nSpellNumber <= 0 Then Exit Function

If Not tabSpells.Fields("Number") = nSpellNumber Then
    tabSpells.Index = "pkSpells"
    tabSpells.Seek "=", nSpellNumber
    If tabSpells.NoMatch Then
        tabSpells.MoveFirst
        Exit Function
    End If
End If

For x = 0 To 9
    If tabSpells.Fields("Abil-" & x) = nAbility Then
        SpellHasAbility = tabSpells.Fields("AbilVal-" & x)
        Exit Function
    End If
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("SpellHasAbility")
Resume out:
End Function

Public Function GetTextblockAction(ByVal nTextblockNumber As Long) As String
On Error GoTo error:

If nTextblockNumber = 0 Then
    GetTextblockAction = "none": Exit Function
End If

tabTBInfo.Index = "pkTBInfo"
tabTBInfo.Seek "=", nTextblockNumber
If tabTBInfo.NoMatch Then
    GetTextblockAction = "none"
    Exit Function
End If
GetTextblockAction = tabTBInfo.Fields("Action")

out:
On Error Resume Next
Exit Function
error:
Call HandleError("GetTextblockAction")
Resume out:
End Function

Public Function GetTextblockCMDS(ByVal nTextblockNumber As Long, Optional ByVal nMaxLength As Integer) As String
Dim x1 As Integer, x2 As Integer, sDecrypted As String

If nTextblockNumber = 0 Then GetTextblockCMDS = "none": Exit Function

tabTBInfo.Index = "pkTBInfo"
tabTBInfo.Seek "=", nTextblockNumber
If tabTBInfo.NoMatch Then
    GetTextblockCMDS = "Textblock " & nTextblockNumber & " not found."
    Exit Function
End If
    
sDecrypted = tabTBInfo.Fields("Action")
If sDecrypted = "" Then
'    If tabTBInfo.Fields("LinkTo") > 0 Then
'        tabTBInfo.Index = "pkTBInfo"
'        tabTBInfo.Seek "=", tabTBInfo.Fields("LinkTo")
'        If tabTBInfo.NoMatch Then
'            GetTextblockCMDS = "none"
'            Exit Function
'        End If
'    Else
        GetTextblockCMDS = "none"
        Exit Function
'    End If
End If

x1 = 1
x1 = InStr(x1, sDecrypted, ":")
If x1 = 0 Then GetTextblockCMDS = "none": Exit Function

GetTextblockCMDS = Mid(sDecrypted, 1, x1 - 1)

x1 = x1 + 1
Do While x1 < Len(sDecrypted)
    x1 = InStr(x1, sDecrypted, Chr(10)) + 1
    If x1 = 1 Then GoTo done:
    
    x2 = InStr(x1, sDecrypted, ":")
    If x2 = 0 Then GoTo done:
    GetTextblockCMDS = GetTextblockCMDS & ", " & Mid(sDecrypted, x1, x2 - x1)
    
    x1 = x2 + 1
Loop


done:
GetTextblockCMDS = Replace(GetTextblockCMDS, "*", "")
GetTextblockCMDS = Replace(GetTextblockCMDS, "|", " OR ")

If nMaxLength > 0 And Len(GetTextblockCMDS) > nMaxLength Then
    GetTextblockCMDS = Left(GetTextblockCMDS, nMaxLength - 1) & "+"
End If

End Function

Public Function GetTextblockTrigger(ByVal nTextblockNumber As Long, ByVal nValue As Long) As String
Dim x1 As Integer
Dim z As Integer, sCommand As String

On Error GoTo error:

If nTextblockNumber = 0 Then GetTextblockTrigger = "none": Exit Function

tabTBInfo.Index = "pkTBInfo"
tabTBInfo.Seek "=", nTextblockNumber
If tabTBInfo.NoMatch Then
    GetTextblockTrigger = "Textblock " & nTextblockNumber & " not found."
    Exit Function
End If

If tabTBInfo.Fields("LinkTo") = nValue Then
    GetTextblockTrigger = "[dialog link]"
    Exit Function
End If

If InStr(1, tabTBInfo.Fields("Action"), "random " & nValue) > 0 Then
    GetTextblockTrigger = "[random " & nValue & "]"
ElseIf InStr(1, tabTBInfo.Fields("Action"), "text " & nValue) > 0 Then
    GetTextblockTrigger = "[text " & nValue & "]"
ElseIf InStr(1, tabTBInfo.Fields("Action"), ":" & nValue) > 0 Then
    
    sCommand = Left(tabTBInfo.Fields("Action"), InStr(1, tabTBInfo.Fields("Action"), ":" & nValue) - 1)
    
    x1 = 1
    Do While x1 < Len(sCommand)
        z = InStr(x1, sCommand, Chr(10))
        If z = 0 Then
            z = x1
            Exit Do
        Else
            x1 = z + 1
        End If
    Loop
    
    sCommand = Right(sCommand, Len(sCommand) - z + 1)
    
    GetTextblockTrigger = "[ask monster " & sCommand & "]"
    
End If

Exit Function
error:
Call HandleError("GetTextblockTrigger")
End Function

Public Function GetTextblockCMDLine(ByVal sCommand As String, Optional ByVal sTextblockData As String, _
    Optional ByVal nTextblockNumber As Long) As String
Dim x1 As Integer, y As Integer
On Error GoTo error:

If nTextblockNumber = 0 And sTextblockData = "" Then
    GetTextblockCMDLine = "unknown": Exit Function
End If

If sTextblockData = "" Then
    tabTBInfo.Index = "pkTBInfo"
    tabTBInfo.Seek "=", nTextblockNumber
    If tabTBInfo.NoMatch Then
        GetTextblockCMDLine = "unknown"
        Exit Function
    End If
    sTextblockData = tabTBInfo.Fields("Action")
End If

'z = slook number
'x1 = end of last command
'x2 = beginning of new command
'slook execution text being searched
'nvalue value of execution
'y = temp position on linebreak

If Not Right(sCommand, 1) = ":" Then sCommand = sCommand & ":"

x1 = InStr(1, sTextblockData, sCommand) 'position x1 at command
If x1 = 0 Then GetTextblockCMDLine = "none": Exit Function
x1 = x1 + Len(sCommand)

y = InStr(x1, sTextblockData, Chr(10))
If y = 0 Then y = Len(sTextblockData)

GetTextblockCMDLine = Mid(sTextblockData, x1, y - x1)
GetTextblockCMDLine = Replace(GetTextblockCMDLine, "*", "")
GetTextblockCMDLine = Replace(GetTextblockCMDLine, "|", " OR ")

Exit Function
error:
Call HandleError("GetTextblockCMDLine")
End Function

Public Function GetTextblockCMDText(ByVal sCommand As String, Optional ByVal sTextblockData As String, _
    Optional ByVal nTextblockNumber As Long) As String
Dim x1 As Integer, sLine As String
On Error GoTo error:

If nTextblockNumber = 0 And sTextblockData = "" Then
    GetTextblockCMDText = "": Exit Function
End If

If sTextblockData = "" Then
    tabTBInfo.Index = "pkTBInfo"
    tabTBInfo.Seek "=", nTextblockNumber
    If tabTBInfo.NoMatch Then
        GetTextblockCMDText = ""
        Exit Function
    End If
    sTextblockData = tabTBInfo.Fields("Action")
End If


x1 = InStr(1, sTextblockData, sCommand) 'position x1 at command
If x1 = 0 Then GetTextblockCMDText = "": Exit Function

sLine = Mid(sTextblockData, 1, x1)

Do While InStr(1, sLine, Chr(10)) > 0
    sLine = Mid(sLine, InStr(1, sLine, Chr(10)) + 1)
Loop
If InStr(1, sLine, ":") > 0 Then
    sLine = Left(sLine, InStr(1, sLine, ":") - 1)
End If

GetTextblockCMDText = sLine
GetTextblockCMDText = Replace(GetTextblockCMDText, "*", "")
GetTextblockCMDText = Replace(GetTextblockCMDText, "|", " OR ")
    
Exit Function
error:
Call HandleError("GetTextblockCMDText")
End Function


Public Sub GetChestItems(ByRef nChestArray() As Currency, ByVal nTBNumber As Long, _
    ByRef nNest As Long, Optional ByVal nPercentMod As Currency)
Dim sData As String, nDataPos As Long, x As Long, y As Long
Dim nPer1 As Long, nPer2 As Long, sLine As String, nValue As Long, nPercent As Currency
Dim nItemArray() As Currency
On Error GoTo error:

tabTBInfo.Index = "pkTBinfo"
tabTBInfo.Seek "=", nTBNumber
If tabTBInfo.NoMatch Then Exit Sub
sData = LCase(tabTBInfo.Fields("Action"))
If sData = Chr(0) Then Exit Sub
nDataPos = 1

If nNest > 5 Then Exit Sub
nNest = nNest + 1

If nPercentMod <= 0 Then nPercentMod = 1
ReDim nItemArray(1 To 2, 0) '1=number, 2=percent

'first we collect all the items and total their %'s
Do While nDataPos < Len(sData)
    x = InStr(nDataPos, sData, ":")
    If x > nDataPos Then
        nPer1 = Val(Mid(sData, nDataPos, x - nDataPos))
        nPercent = (nPer1 - nPer2) / 100
        nPer2 = nPer1
        
        nDataPos = x + 1
        'nNest = nNest + 1
        
        x = InStr(nDataPos, sData, Chr(10))
        If x <= 0 Then x = Len(sData)
        sLine = LCase(Mid(sData, nDataPos, x - nDataPos))
        nDataPos = x
        
        y = 1
check_give_again:
        y = InStr(y, sLine, "giveitem ")
        If y > 0 Then
            nValue = ExtractValueFromString(Mid(sLine, y), "giveitem ")
            
            For x = 0 To UBound(nItemArray(), 2)
                If nItemArray(1, x) = nValue Then
                    nItemArray(2, x) = nItemArray(2, x) + nPercent
                    x = -1
                    Exit For
                End If
            Next x
            If x >= 0 Then
                x = UBound(nItemArray(), 2) + 1
                ReDim Preserve nItemArray(1 To 2, x)
                nItemArray(1, x) = nValue
                nItemArray(2, x) = nPercent
            End If
            
            y = y + 1
            GoTo check_give_again:
        End If
        
        y = 1
check_random_again:
        y = InStr(y, sLine, "random ")
        If y > 0 Then
            
            nValue = ExtractValueFromString(Mid(sLine, y), "random ")
            If nValue > 0 Then
                Call GetChestItems(nChestArray(), nValue, nNest, (nPercent * nPercentMod))
            End If
            
            y = y + 1
            GoTo check_random_again:
        End If
        
'''''        If InStr(1, sLine, "giveitem ") > 0 Then
'''''            nValue = ExtractValueFromString(sLine, "item ")
'''''
'''''            For x = 0 To UBound(nItemArray(), 2)
'''''                If nItemArray(1, x) = nValue Then
'''''                    nItemArray(2, x) = nItemArray(2, x) + nPercent
'''''                    x = -1
'''''                    Exit For
'''''                End If
'''''            Next x
'''''            If x >= 0 Then
'''''                x = UBound(nItemArray(), 2) + 1
'''''                ReDim Preserve nItemArray(1 To 2, x)
'''''                nItemArray(1, x) = nValue
'''''                nItemArray(2, x) = nPercent
'''''            End If
'''''
'''''        ElseIf InStr(1, sLine, "random ") > 0 Then
'''''            nValue = ExtractValueFromString(sLine, "random ")
'''''            If nValue > 0 Then
'''''                Call GetChestItems(nChestArray(), nValue, nNest, (nPercent * nPercentMod))
'''''            End If
'''''        End If
    Else
        nDataPos = nDataPos + 1
    End If
Loop

'then we put the collected items into the chest array
'...this is actually sort of unecessary i found out afterwards
For y = 0 To UBound(nItemArray(), 2)
    If nItemArray(1, y) > 0 Then
        nPercent = nItemArray(2, y)
        
        For x = 0 To UBound(nChestArray(), 2)
            If nChestArray(1, x) = nItemArray(1, y) Then
                'If nChestArray(3, x) = 0 Then nChestArray(3, x) = 1
                nChestArray(2, x) = nChestArray(2, x) + _
                    (nChestArray(3, x) * nPercent * nPercentMod)
                nChestArray(3, x) = nChestArray(3, x) * (1 - nPercent)
                x = -1
                Exit For
            End If
        Next x
        If x >= 0 Then
            x = UBound(nChestArray(), 2) + 1
            ReDim Preserve nChestArray(1 To 3, x)
            nChestArray(1, x) = nItemArray(1, y)
            nChestArray(2, x) = nPercent * nPercentMod
            nChestArray(3, x) = 1 - nChestArray(2, x)
        End If
        
    End If
Next y

nNest = nNest - 1

Erase nItemArray()

Exit Sub
error:
Call HandleError("GetChestItems-#" & nTBNumber)
Erase nItemArray()
End Sub

Public Function CalculateMonsterItemBonuses(nMonster As Long, nAbilities As Variant) As Integer
Dim x As Integer, y As Integer, nTest As Integer
On Error GoTo error:

If Not IsDimmed(nAbilities) Then Exit Function

If tabMonsters.Fields("Number") <> nMonster Then
    tabMonsters.Index = "pkMonsters"
    tabMonsters.Seek "=", nMonster
    If tabMonsters.NoMatch = True Then
        tabMonsters.MoveFirst
        Exit Function
    End If
End If

If tabMonsters.Fields("Weapon") > 0 Then
    If GetItemLimit(tabMonsters.Fields("Weapon")) = 0 Then
        For y = LBound(nAbilities) To UBound(nAbilities)
            nTest = ItemHasAbility(tabMonsters.Fields("Weapon"), nAbilities(y))
            If nTest > 0 Then
                CalculateMonsterItemBonuses = CalculateMonsterItemBonuses + nTest
            End If
        Next y
    End If
End If

For x = 0 To 9
    If tabMonsters.Fields("DropItem-" & x) > 0 Then
        If GetItemLimit(tabMonsters.Fields("DropItem-" & x)) = 0 Then
            For y = LBound(nAbilities) To UBound(nAbilities)
                nTest = ItemHasAbility(tabMonsters.Fields("DropItem-" & x), nAbilities(y))
                If nTest > 0 Then
                    If tabMonsters.Fields("DropItem%-" & x) > 100 Then tabMonsters.Fields("DropItem%-" & x) = 100
                    CalculateMonsterItemBonuses = CalculateMonsterItemBonuses + (nTest * (tabMonsters.Fields("DropItem%-" & x) / 100))
                End If
            Next y
        End If
    End If
Next x

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalculateMonsterItemBonuses")
Resume out:
End Function

Public Sub PopulateMonsterDataToAttackSim(ByVal nMonster As Long, ByRef clsMonAtkSim As clsMonsterAttackSim)
On Error GoTo error:
Dim x As Integer, y As Integer
Dim nPercent As Integer, sTemp As String, nTest As Integer
Dim nDamageArr As Variant, nAccyArr As Variant
Dim nItemDamageBonus As Integer, nItemAccyBonus As Integer

On Error Resume Next
If tabMonsters.Fields("Number") <> nMonster Then
    tabMonsters.Index = "pkMonsters"
    tabMonsters.Seek "=", nMonster
    If tabMonsters.NoMatch = True Then
        tabMonsters.MoveFirst
        Exit Sub
    End If
End If
On Error GoTo error:

If nNMRVer >= 1.71 Then
    clsMonAtkSim.nEnergyPerRound = tabMonsters.Fields("Energy")
Else
    clsMonAtkSim.nEnergyPerRound = 1000
End If

nDamageArr = Array(4) '4=max damage
nAccyArr = Array(22, 105, 106) '22, 105, 106 = accuracy

nItemDamageBonus = CalculateMonsterItemBonuses(nMonster, nDamageArr)
nItemAccyBonus = CalculateMonsterItemBonuses(nMonster, nAccyArr)

For x = 0 To 4
    sTemp = ""
    If tabMonsters.Fields("AttType-" & x) > 0 And tabMonsters.Fields("AttType-" & x) < 4 Then
        If nNMRVer >= 1.8 Then
            sTemp = tabMonsters.Fields("AttName-" & x)
        Else
            If tabMonsters.Fields("AttType-" & x) = 2 And tabMonsters.Fields("AttAcc-" & x) > 0 Then
                sTemp = GetSpellName(tabMonsters.Fields("AttAcc-" & x), True)
            End If
            If sTemp = "" Or sTemp = "None" Then sTemp = "Attack " & (x + 1)
        End If
        clsMonAtkSim.sAtkName(x) = Trim(sTemp)
        clsMonAtkSim.nAtkType(x) = tabMonsters.Fields("AttType-" & x)
        clsMonAtkSim.nAtkEnergy(x) = tabMonsters.Fields("AttEnergy-" & x)
        clsMonAtkSim.nAtkChance(x) = tabMonsters.Fields("Att%-" & x)
        
        If tabMonsters.Fields("AttType-" & x) = 2 Then 'spell
            
            tabSpells.Index = "pkSpells"
            tabSpells.Seek "=", tabMonsters.Fields("AttAcc-" & x)
            If tabSpells.NoMatch = True Then
                tabSpells.MoveFirst
                GoTo next_attack_slot:
            Else
                If tabSpells.Fields("Targets") = 12 Then
                    If GetSpellDuration(tabMonsters.Fields("AttAcc-" & x), tabMonsters.Fields("AttMax-" & x), True) = 0 Then
                        nTest = SpellHasAbility(tabMonsters.Fields("AttAcc-" & x), 1) '1=damage
                        If nTest > -1 Then
                            'MsgBox "Attack #" & (x + 1) & " (" & txtAtkName(x).Text & ") has an area attack spell in a regular attack slot using ability 1 (damage) instead of 17 (damage-MR). " _
                                & "This is an error and MMUD will not cast this.  Area attack spells must use ability 17 (or possibly 8-drain?).  The min/max damage and energy cost has been zero'd out for the sim to reflect the game.", vbExclamation
                            clsMonAtkSim.nAtkDuration(x) = 0
                            clsMonAtkSim.nAtkMin(x) = 0
                            clsMonAtkSim.nAtkMax(x) = 0
                            clsMonAtkSim.nAtkEnergy(x) = 0
                            GoTo next_attack_slot:
                        End If
                    End If
                End If
                
                If nNMRVer >= 1.8 Then clsMonAtkSim.nAtkResist(x) = tabSpells.Fields("TypeOfResists")
                
                clsMonAtkSim.nAtkDuration(x) = GetSpellDuration(tabMonsters.Fields("AttAcc-" & x), tabMonsters.Fields("AttMax-" & x), True)
                clsMonAtkSim.nAtkMin(x) = 0
                clsMonAtkSim.nAtkMax(x) = 0
                
                nTest = SpellHasAbility(tabMonsters.Fields("AttAcc-" & x), 1) '1=damage
                If nTest >= 0 Then
                    clsMonAtkSim.nAtkMRdmgResist(x) = 0 'NO MR resist
                    If nTest > 0 Then
                        clsMonAtkSim.nAtkMin(x) = nTest
                        clsMonAtkSim.nAtkMax(x) = nTest
                    Else
                        clsMonAtkSim.nAtkMin(x) = GetSpellMinDamage(tabMonsters.Fields("AttAcc-" & x), tabMonsters.Fields("AttMax-" & x), -1, True)
                        clsMonAtkSim.nAtkMax(x) = GetSpellMaxDamage(tabMonsters.Fields("AttAcc-" & x), tabMonsters.Fields("AttMax-" & x), -1, True)
                    End If
                End If
                
                nTest = SpellHasAbility(tabMonsters.Fields("AttAcc-" & x), 17) '17=damage
                If nTest >= 0 Then
                    clsMonAtkSim.nAtkMRdmgResist(x) = 1 'MR resist
                    If nTest > 0 Then
                        clsMonAtkSim.nAtkMin(x) = nTest
                        clsMonAtkSim.nAtkMax(x) = nTest
                    Else
                        clsMonAtkSim.nAtkMin(x) = GetSpellMinDamage(tabMonsters.Fields("AttAcc-" & x), tabMonsters.Fields("AttMax-" & x), -1, True)
                        clsMonAtkSim.nAtkMax(x) = GetSpellMaxDamage(tabMonsters.Fields("AttAcc-" & x), tabMonsters.Fields("AttMax-" & x), -1, True)
                    End If
                End If
                
                nTest = SpellHasAbility(tabMonsters.Fields("AttAcc-" & x), 8) '8=drain
                If nTest >= 0 Then
                    clsMonAtkSim.nAtkMRdmgResist(x) = 0 'NO MR resist
                    If nTest > 0 Then
                        clsMonAtkSim.nAtkMin(x) = nTest
                        clsMonAtkSim.nAtkMax(x) = nTest
                    Else
                        clsMonAtkSim.nAtkMin(x) = GetSpellMinDamage(tabMonsters.Fields("AttAcc-" & x), tabMonsters.Fields("AttMax-" & x), -1, True)
                        clsMonAtkSim.nAtkMax(x) = GetSpellMaxDamage(tabMonsters.Fields("AttAcc-" & x), tabMonsters.Fields("AttMax-" & x), -1, True)
                    End If
                End If
            End If
            clsMonAtkSim.nAtkSuccess(x) = tabMonsters.Fields("AttMin-" & x)
        Else
            clsMonAtkSim.nAtkMin(x) = tabMonsters.Fields("AttMin-" & x) + nItemDamageBonus
            clsMonAtkSim.nAtkMax(x) = tabMonsters.Fields("AttMax-" & x) + nItemDamageBonus
            clsMonAtkSim.nAtkSuccess(x) = tabMonsters.Fields("AttAcc-" & x) + nItemAccyBonus
            If tabMonsters.Fields("AttHitSpell-" & x) > 0 Then
                
                tabSpells.Index = "pkSpells"
                tabSpells.Seek "=", tabMonsters.Fields("AttHitSpell-" & x)
                If tabSpells.NoMatch = True Then
                    tabSpells.MoveFirst
                    GoTo next_attack_slot:
                Else
                    If nNMRVer >= 1.8 Then clsMonAtkSim.nAtkResist(x) = tabSpells.Fields("TypeOfResists")
                    clsMonAtkSim.sAtkHitSpellName(x) = tabSpells.Fields("Name")
                    clsMonAtkSim.nAtkDuration(x) = GetSpellDuration(tabMonsters.Fields("AttHitSpell-" & x))
                    
                    If SpellHasAbility(tabMonsters.Fields("AttHitSpell-" & x), 1) >= 0 Then
                        clsMonAtkSim.nAtkMRdmgResist(x) = 0
                        clsMonAtkSim.nAtkHitSpellMin(x) = GetSpellMinDamage(tabMonsters.Fields("AttHitSpell-" & x))
                        clsMonAtkSim.nAtkHitSpellMax(x) = GetSpellMaxDamage(tabMonsters.Fields("AttHitSpell-" & x))
                        
                    ElseIf SpellHasAbility(tabMonsters.Fields("AttHitSpell-" & x), 17) >= 0 Then
                        clsMonAtkSim.nAtkMRdmgResist(x) = 1
                        clsMonAtkSim.nAtkHitSpellMin(x) = GetSpellMinDamage(tabMonsters.Fields("AttHitSpell-" & x))
                        clsMonAtkSim.nAtkHitSpellMax(x) = GetSpellMaxDamage(tabMonsters.Fields("AttHitSpell-" & x))
                        
                    Else
                        clsMonAtkSim.nAtkHitSpellMin(x) = 0
                        clsMonAtkSim.nAtkHitSpellMax(x) = 0
                    End If
                End If
            End If
        End If
    End If
next_attack_slot:
Next x

nPercent = 0
For x = 0 To 4
    If tabMonsters.Fields("MidSpell-" & x) > 0 Then
        tabSpells.Index = "pkSpells"
        tabSpells.Seek "=", tabMonsters.Fields("MidSpell-" & x)
        If tabSpells.NoMatch = True Then
            tabSpells.MoveFirst
            'GoTo next_attack_slot:
        Else
            clsMonAtkSim.sBetweenRoundName(x) = tabSpells.Fields("Name")
            If nNMRVer >= 1.8 Then clsMonAtkSim.nBetweenRoundResistType(x) = tabSpells.Fields("TypeOfResists")
            clsMonAtkSim.nBetweenRoundChance(x) = tabMonsters.Fields("MidSpell%-" & x)
            clsMonAtkSim.nBetweenRoundDuration(x) = GetSpellDuration(tabMonsters.Fields("MidSpell-" & x), tabMonsters.Fields("MidSpellLVL-" & x), True)
            
            nTest = SpellHasAbility(tabMonsters.Fields("MidSpell-" & x), 1) '1=damage
            If nTest >= 0 Then
                clsMonAtkSim.nBetweenRoundResistDmgMR(x) = 0 'NO MR resist
                If nTest > 0 Then
                    clsMonAtkSim.nBetweenRoundMin(x) = nTest
                    clsMonAtkSim.nBetweenRoundMax(x) = nTest
                Else
                    clsMonAtkSim.nBetweenRoundMin(x) = GetSpellMinDamage(tabMonsters.Fields("MidSpell-" & x), tabMonsters.Fields("MidSpellLVL-" & x), -1, True)
                    clsMonAtkSim.nBetweenRoundMax(x) = GetSpellMaxDamage(tabMonsters.Fields("MidSpell-" & x), tabMonsters.Fields("MidSpellLVL-" & x), -1, True)
                End If
            End If
            
            nTest = SpellHasAbility(tabMonsters.Fields("MidSpell-" & x), 17) '17=damage-mr
            If nTest >= 0 Then
                clsMonAtkSim.nBetweenRoundResistDmgMR(x) = 1 'MR resist
                If nTest > 0 Then
                    clsMonAtkSim.nBetweenRoundMin(x) = nTest
                    clsMonAtkSim.nBetweenRoundMax(x) = nTest
                Else
                    clsMonAtkSim.nBetweenRoundMin(x) = GetSpellMinDamage(tabMonsters.Fields("MidSpell-" & x), tabMonsters.Fields("MidSpellLVL-" & x), -1, True)
                    clsMonAtkSim.nBetweenRoundMax(x) = GetSpellMaxDamage(tabMonsters.Fields("MidSpell-" & x), tabMonsters.Fields("MidSpellLVL-" & x), -1, True)
                End If
            End If
            
            nTest = SpellHasAbility(tabMonsters.Fields("MidSpell-" & x), 8) '8=drain
            If nTest >= 0 Then
                clsMonAtkSim.nBetweenRoundResistDmgMR(x) = 0 'NO MR resist
                If nTest > 0 Then
                    clsMonAtkSim.nBetweenRoundMin(x) = nTest
                    clsMonAtkSim.nBetweenRoundMax(x) = nTest
                Else
                    clsMonAtkSim.nBetweenRoundMin(x) = GetSpellMinDamage(tabMonsters.Fields("MidSpell-" & x), tabMonsters.Fields("MidSpellLVL-" & x), -1, True)
                    clsMonAtkSim.nBetweenRoundMax(x) = GetSpellMaxDamage(tabMonsters.Fields("MidSpell-" & x), tabMonsters.Fields("MidSpellLVL-" & x), -1, True)
                End If
            End If
        End If
    End If
Next x

For x = 0 To 4
    If Len(clsMonAtkSim.sAtkName(x)) > 0 Then
        For y = 0 To 4
            If y <> x And clsMonAtkSim.sAtkName(x) = clsMonAtkSim.sAtkName(y) Then
                clsMonAtkSim.sAtkName(x) = Trim(clsMonAtkSim.sAtkName(x)) & (x + 1)
                clsMonAtkSim.sAtkName(y) = Trim(clsMonAtkSim.sAtkName(y)) & (y + 1)
            End If
        Next y
    End If
Next x

out:
On Error Resume Next
Exit Sub
error:
Call HandleError("PopulateMonsterDataToAttackSim")
Resume out:
End Sub

Public Function CalculateMonsterAvgDmg(ByVal nMonster As Long, Optional nNumRounds As Long = 500) As Variant()
On Error GoTo error:

Call clsMonAtkSim.ResetValues
clsMonAtkSim.bUseCPU = True
clsMonAtkSim.nCombatLogMaxRounds = 0
clsMonAtkSim.nNumberOfRounds = nNumRounds 'IIf(nNumRounds <> 0, nNumRounds, 500)
clsMonAtkSim.nUserMR = 50
clsMonAtkSim.bDynamicCalc = False
clsMonAtkSim.nDynamicCalcDifference = 0.001

Call PopulateMonsterDataToAttackSim(nMonster, clsMonAtkSim)

If clsMonAtkSim.nNumberOfRounds > 0 Then clsMonAtkSim.RunSim

ReDim nReturn(1)
nReturn(0) = clsMonAtkSim.nAverageDamage
nReturn(1) = clsMonAtkSim.GetMaxDamage

CalculateMonsterAvgDmg = nReturn

out:
On Error Resume Next
Exit Function
error:
Call HandleError("CalculateMonsterAvgDmg")
Resume out:
End Function

Public Function IsDimmed(Arr As Variant) As Boolean
On Error GoTo ReturnFalse
  IsDimmed = UBound(Arr) >= LBound(Arr)
ReturnFalse:
End Function
