# MMUD-Explorer  

MMUD Explorer is a database viewer for the game MajorMUD(r) created by syntax53. It has a unique comparing feature which allows you to easily compare weapons, armour, and spells. It also has a graphical room explorer in which you can 'walk' around the realm. Other features include an inventory calculator, exp calculator, explorers for monsters/shops/weapons/armour/spells/items/races/classes, saving/loading characters, and copying data to and from the clipboard.  More info may be found here: http://www.mudinfo.net/viewforum.php?f=34  

v1.8 (05/??/2016) - By Syntax  
------------------------------------------  
General:  
-NEW: added recent database list for fast switching  
-UP: added the option to only copy the name of something when right-clicking and copying  
-UP: increased the minimum width and heigh of the main window  
-UP: re-organized global filter to allow you to filter by level and alignment without having a class selected  
-UP: shops will now show the appropriate coin type value next to the copper value when buying, selling, training  
-UP: checking "simple" on the textblock results window will no longer hide "minlevel #"  

Character:  
-UP: load character form will now be displayed when loading from recent list as well  
-UP: made using the second wrist slot default (for new installs)  
-UP: added MR value corresponding resistance values  
-UP: added AC value to show miss rate vs accuracy values  
-FIX: pasting of equipped eye, face, and back items
-FIX: second alignment checkbox not saving on character tab  

Spells:  
-NEW: new field (from NMR v1.8+ exports) for displaying if spell is resistible by no one, anyone, and anti-magic  
-NEW: damage/mana efficiency column on spells  
-UP: spells restricted to classes via learning method now filtered (requires database created with NMR v1.7+)  
-UP: level used for calculating spell damage no longer minimized to required level (monsters don't obey req. level)  
-UP: if a spell executes a textblock and that textblock only casts spells, the spells will be displayed in the detail window  

Items:  
-NEW: filter weapons and armor by ability value  
-UP: equipped items will now be highlighted on the weapons and armour tabs  
-UP: added an "ability" column which will show the filtered ability value (won't show on compare window though)  
-UP: added a crits column to weapons and armor  
-UP: added an "average damage * 5" column to weapons  
-UP: item/key references for rooms and textblocks expanded (requires database created with NMR v1.7+)  
-UP: find best will now take AC/Enc ratio into account when stats are equal  
-UP: hardcoded item #s 68 & 100 (dragger & quarterstaff) as usable by "limited" weapon classes as is in the DLL  
-UP: added a non-crit average damage to swing calc
-UP: increased magic level dropdown on weapons from 6+ to 10+  
-FIX: changed the upper default weapon speed limit to 999999 (mudrev weapons, in particular)  

Monsters:  
-NEW: new field (from NMR v1.7.1+ exports) for reporting and evaluating monster energy & attacks accurately  
-NEW: columns on monster list: added "# lairs (% average lairs)", (and for NMR v1.8+ exports): average damage and "scripting value" which is the number of lairs for the monster compared to the average lairs for the database multiplied by (EXP/(DMG+HP))  
-UP: added coloring to health  
-FIX: min/max spell damage was sometimes incorrect if the cast level was outside the spell's min/max level caps  

Rooms:  
-NEW: Expanded map to 50x50! Added options for 30x50 and 50x50  
-NEW: Added options to external map to mark shops or room spells  
-UP: expanded map size on rooms tab as well to 30x23 (was 20x20)
-UP: Added option to external map to permit the drawing duplicate rooms  
-FIX: Fixed presets. Consider resetting them!  Removed the "feature" that stores them specific to database  


v1.70 (02/18/2008) - By Ghaleon  
------------------------------------------  
General:  
-FIX: error when pasting char info into character tab  
-FIX: remove on equip tab removes face/eyes  
-FIX: weapon dmg/spd should sort correctly  

v1.69 (10/02/2007) - By Ghaleon  
------------------------------------------  
Weapons  
-UP: separated damage into min/max so it is sortable  

Compare  
-UP: separated damage into min/max so it is sortable  

Equipment  
-NEW: added eyes & face  
-NEW: added mystic powers to the 'find best'  

General  
-UP: added eyes & face to the ini save/load function  

v1.68 (??)  
------------------------------------------   
General:  
-FIX: repeated prompts about app running when switching dats  

Spells:  
-UP: RemoveSpells will now display after the number of rounds  
-FIX: reported DR amount in spell formula when value is scaled  
-FIX: endcast'ed spell display when not testing against level  


v1.67 (07/19/2004)  
------------------------------------------   
General:  
-UP: spells should load a little faster  
-UP: items with +acc abilities will show in the column  
-UP: items wont show level, acc, or bs accu in detail box  
-FIX: error when closing program with bs calc open  
-FIX: bs formula... again... hopefully  


v1.66 (07/18/2004)  
------------------------------------------   
General:  
-FIX: prompts for character save after removing a filter  
-FIX: find best on wrist and finger slots  


v1.65 (07/17/2004)  
------------------------------------------   
General:  
-UP: wont prompt for save unless something has been changed  
-UP: added quest box to select 2nd alignment quest bonuses  
-UP: item/spell reference lists are now sorted  
-UP: expanded width/height of a lot of the combo boxes  
-UP: improved column sorting  
-FIX: doubling of compare lists when loading/reverting/closing  
-FIX: pasting recognition of (Two Handed)  
-FIX: tweaked the BS Calculator formula  
-FIX: number only fields will allow control+c/v  
-FIX: CP calculation issues  
-FIX: tab order  
-FIX: monster compare right click menu  
-FIX: textblocks wouldn't show repeated count for last command  
-FIX: notepad annoyances with focus and shortcut keys  

Equipment:  
-UP: added some find bests and put them in sub-categories  
-UP: equiped items now show more stats in tooltip  
-FIX: add all to compare on equipment page  
-FIX: typing/selection issues on equipment & char blesses  

Items:  
-NEW: sundry tab shows chest contents in a nice list (button)  
-UP: items from NPCs will show the NPC instead of the textblock  
-UP: items from chests will now show the chest too  
-UP: scrolls will now show a reference to the spell they teach  
-FIX: if you unselect item/armour types, they'll filter out +  

Rooms:  
-UP: now show placed items  
-UP: now show the key/item name within the tooltip  
-UP: now shows whether or not doors are bashable  

Spells:  
-UP: spell formulas will update when turning off global filter  
-UP: spells will show the NPC learned from instead of textblock  
-UP: wont show (@lvl X) if there is no need for it  

+I had actually intended for it to work the way it   
was working (only filtering out item types if the   
global filter was UN-selected), it was just a miserable   
failure...  


v1.61 (05/25/2004)  
------------------------------------------   
General:  
-NEW: backstab calculator  
-NEW: notepad tool for pasting temporary info  
-NEW: added an option to show character name in title  
-NEW: recent character list on file menu  
-UP: navigation buttons will now span (option to disable)  
-UP: paste character screen is now sizable  
-UP: will now prompt for save on character load  
-UP: now saves window position  
-FIX: auto-complete issues on equipment & bless boxes  
-FIX: find room progress not showing on external map  
-FIX: fixed some room lines not connecting/drawing straight  

Items:  
-NEW: Encumbrance/AC ratio column for armour   
-NEW: Speed/Damage ratio column for weapons  
-UP: added partial support for future face/eye slots  
-FIX: class restricted/ok'd weapons & armour not filtering ++  
-FIX: worn-on labels when copying inventory to clipboard  
-FIX: recognation of first pasted item when no user gender  

Monsters:  
-NEW: added splitter bar for monsters  
-NEW: added a Exp/HP ratio column  
-UP: moved abilities into normal stat window  
-UP: regen information now shown in normal stat window  
-UP: most bogus "Greet Commands" will be hidden  

Spells:  
-UP: shows # of times casted for spells like magma blast  
-UP: now adds a clickable reference for monster summons  
-UP: bless spells show the spell formula in the tooltip  
-UP: expanded width of bless spell dropdown  

++NOTE: all armour types are now enabled no matter what  
armour type the class is.  however, if the global filter  
is enabled, only items the class can use will be filtered.  


v1.6 (03/22/2004)  
------------------------------------------   
Rooms:  
-NEW: external map viewer! Stays on top of MegaMUD, 30x30  
-UP: moved preset saves to the registry  
-FIX: fixed tooltip display of some hidden exits +  
-FIX: 2+ teleport cmds in one room displayed wrong text  
-NOTE: You will need to recreate any presets you've set  

General:  
-UP: can now press the F key to re-select-all the find field  
-UP: weapon/armour/spell splitters now synced  
-FIX: fixed column size of reference lists on load  
-FIX: fixed color issues with different appearance settings  

Compare:  
-NEW: Added a Monster compare section  

Items:  
-FIX: wasn't showing when an item destroyed on death  

Results Window:  
(For Textblocks ...)  
-UP: will now dig into testskill commands +  
-UP: will now save window position for each window type  
-FIX: some blocks with one command would show as dialog  

Swing Calc:  
-NEW: full true damage calculator (editable fields)  
-UP: more copy to clipboard options  
-UP: will now copy bashing status  

+ Requires updated mmud explorer database  


v1.51 (03/06/2004)  
------------------------------------------   
General:  
-UP: added a "Revert to Saved" menu item  
-FIX: fixed a loading issue with the settings.ini  
-FIX: fixed bug with positioning splitters when maximized  
-FIX: updated installer for full compatibility with Win95+  
-NOTE: thanks again to paxtez for ideas/help on the *'s  

Character Tab:  
-NEW: bless spell calculator *  
-NEW: true damage calculator (also on swing calc) *  
-UP: added a reload cp button to reload from file *  
-UP: CP Range box now shows tooltip of CPs used *  

Rooms:  
-FIX: rooms drew unconnected lines on some screen configs  
-FIX: rooms displayed actions in reverse, now sorted too *  

Results Window:  
(For Textblocks ...)  
-UP: added "expand/collapse branch" right click menus  
-UP: duplicate lines will combine to "Line (x20)"  
-UP: removed "--> Raw" lines (click ? to see alternative)  


v1.5 (03/03/2004)  
------------------------------------------   
General:  
-NEW: swing calculator ++  
-NEW: MUCH better and very cool textblock display/handling  
-NEW: can now resize the windows on item/spell tabs  
-UP: better column sort handling  
-UP: ability to copy the exp calc results to the clipboard  
-UP: you can now copy multiple listview lines to the clipboard  
-UP: will now prompt on adding a dupe record to the compares *  
-UP: item/spells will now prompt for unfilter when not found *  
-UP: A LOT of code optimization  
-UP: class/race tabs were combined into a "Descent" tab  
-UP: character files now show that they're loaded and such  
-FIX: "F" keys wont trigger when shift/control/alt is pressed  
-NOTE: Thanks to "Paxtez" for submitting ideas on all the *'s  
-NOTE: Huge thanks to Locke for the formulas on all the ++'s  

Settings:  
-NEW: option to hide room numbers when looking up record names  
-NEW: option to make exp calc/result windows not "on top"  
-NEW: option to swap left/right mouse button functions for map  
-NEW: options to auto load/save characters  

Character Tab (NEW):  
-NEW: CP calculator  
-NEW: HP & HP Regen Calculators ++  
-NEW: Mana & Mana Regen Calculators ++  
-NEW: spellcasting calculator ++  
-NEW: picklocks calculator ++  
-NEW: checkboxes for completed quests  
-NEW: copy stats to clipboard  
-NEW: copy CP info to clipboard  

Shops:  
-NEW: cost shown for buying/selling/training  
-UP: sorting by item cost now sorts by the copper value *  
-UP: when showing a shop reference, the room name will be shown  
-FIX: fixed some rounding issues on item cost  

Equipment (WAS 'Inventory')  
-NEW: next best button to find next item of equal/lesser value  
-NEW: added 'hold item' checkboxes  
-NEW: added an "additional items" weight field  
-NEW: add all items selected to compare  
-NEW: empty lists button (see help button next to it)  
-NEW: checkbox to hide character specific stats  
-UP: Find Best won't choose a 2-Hand weapon and an off-hand  
-UP: Find Best wont select a shield for 1H weapon classes  
-UP: goto buttons now prompt for goto or compare *  
-UP: encum status now shows when you get light/med/heavy *  
-UP: worn labels will turn red when shield & 2H weapon selected *  
-FIX: copy to clipboard  
-FIX: find best works much faster now  

Rooms:  
-NEW: lair rooms are now shown with what regens in them  
-NEW: teleport rooms are now supported and shows the commands  
-NEW: full display of room commands and their scripts  
-NEW: presets are now customizable, and there are 50  
-NEW: button to add/copy megamud codes on path recorder  
-NEW: button to search for a partial/full room name  
-NEW: button to go back to last room  
-NEW: button to enable manual typing into move box  
-UP: now shows picklock values and trap damage  
-UP: map won't flicker when changing rooms anymore  
-FIX: path steps recorded from map have been reversed (duh!) *  

Spells:  
-NEW: filter spells by containing a certain ability  
-UP: Spells with textblocks now allow you to view the commands  
-UP: A LOT of work to show spell formulas more accurately  
-UP: RemoveSpells combined to one line  
-UP: switched where/how the spell tab displays the formula  

Items:  
-UP: items show spell references in the location boxes  
-UP: option to jump to compare/equip window when adding an item  
-FIX: fixed AC not showing decimals on armour tab  

Monsters:  
-NEW: greet textblock as well as the script for the commands  
-UP: added a magical column for monsters  
-UP: mons guards combines to one line  
-FIX: mons showed wrong stats when a spell summoned other mons  


v1.1 (12/17/2003)  
------------------------------------------   
-FIX: AC and DR calculation issues on inventory  
-UP: Added some more "Find Best..." items  
-UP: Added the ability to change the inventory fonts  


v1.0 (11/26/2003)  
------------------------------------------   
-NEW: MMUD Explorer Released!  
-NOTE: thanks to locke for the exp calculator function  