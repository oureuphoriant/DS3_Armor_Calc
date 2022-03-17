import csv
import os.path
import PySimpleGUI as sg
from operator import attrgetter

if not os.path.exists("DS3_Armor_Calc.csv"):
    # csvBackup = [
    #     ["active","name","type","weight","standard","strike","slash","thrust","magic","fire","lightning","dark","bleed","poison","frost","curse","poise"],
    #     ["N","Alva Armor","chest","","","","","","","","","","","","","",""],
    #     ["N","Alva Gauntlets","hands","","","","","","","","","","","","","",""],
    #     ["N","Alva Helm","head","","","","","","","","","","","","","",""],
    #     ["N","Alva Leggings","legs","","","","","","","","","","","","","",""],
    #     ["Y","Antiquated Dress","chest",3.1,3.2,3.6,3.2,3.2,12.3,13.9,11.8,13.2,21,38,27,86,1.2],
    #     ["Y","Antiquated Gloves","hands",1.1,0.9,1,0.9,0.9,3.3,3.7,3.2,3.6,8,13,10,29,0.6],
    #     ["Y","Antiquated Plain Garb","chest",3.1,3.7,3.1,2.7,2.4,12.7,11.1,12.2,12.3,30,47,25,60,0.9],
    #     ["Y","Antiquated Skirt","legs",2.1,1.6,1.9,1.6,1.4,6.7,7.5,6.4,7,13,23,16,54,0.4],
    #     ["Y","Archdeacon Holy Garb","chest",4.2,3.7,6.8,4.4,4.4,12.4,11.5,12.4,13.7,18,39,37,83,3.4],
    #     ["Y","Archdeacon Skirt","legs",2.6,2.4,4.2,2.9,2.9,7.7,7.2,7.7,8.5,14,27,26,55,2.7],
    #     ["Y","Archdeacon White Crown","head",1.8,1.6,2.7,1.9,1.9,5,4.6,5,5.4,11,19,19,38,1.6],
    #     ["Y","Aristocrat's Mask","head",5.7,6,2.7,5.8,5.5,4.1,4.3,2.5,3.9,25,17,17,28,6.8],
    #     ["Y","Armor of the Sun","chest",8.6,11.2,8.3,11.2,11.2,10.5,10.5,9.8,10.5,34,28,31,26,11.3],
    #     ["Y","Assassin Armor","chest",6.9,8.6,10,8.6,9.3,10.4,11.4,10.4,9,40,62,47,52,4.8],
    #     ["Y","Assassin Gloves","hands",2,2,1.2,2.3,1.8,2.6,2.7,1.9,1.5,12,20,13,16,0.4],
    #     ["Y","Assassin Hood","head",1.8,1.9,2.7,1.9,2.2,3.3,4.3,3.3,2.5,17,27,20,20,1],
    #     ["Y","Assassin Trousers","legs",4.3,5.2,3.8,6,4.8,6.6,6.9,5.2,4.3,29,43,32,38,2.1],
    #     ["Y","Armor of Thorns","chest",8.5,11,8.1,14.2,11.5,7.3,7.3,7.3,5.9,66,27,33,22,7.5],
    #     ["N","Billed Mask","head","","","","","","","","","","","","","",""],
    #     ["N","Black Dress","chest","","","","","","","","","","","","","",""],
    #     ["N","Black Gauntlets","hands","","","","","","","","","","","","","",""],
    #     ["Y","Black Hand Armor","chest",7.8,9,8.1,11.5,9.7,12.6,12.8,10.4,9,46,67,51,58,4.9],
    #     ["Y","Black Hand Hat","head",2.5,2.6,2.1,3.8,2.9,4.2,4.4,3.3,2.7,19,30,21,25,1.1],
    #     ["Y","Black Iron Armor","chest",15.8,18.1,12.2,18.1,18.1,12.2,15.8,11.5,12.2,65,51,57,35,20.9],
    #     ["Y","Black Iron Gauntlets","hands",5.1,4.6,2.9,4.6,4.4,2.8,3.8,2.8,3.1,24,16,20,14,4.4],
    #     ["Y","Black Iron Helm","head",6.5,6.5,4.5,6.9,6.6,4.8,5.9,4.4,4.8,26,13,21,13,7.6],
    #     ["Y","Black Iron Leggings","legs",9.5,10.8,7,10.8,10.3,7.5,9.3,6.9,7.5,42,32,35,24,13.4],
    #     ["Y","Black Knight Armor","chest",13.4,15.8,12.2,14.4,13.8,11,14.1,10.5,12,58,42,39,39,15.8],
    #     ["N","Black Knight Gauntlets","hands","","","","","","","","","","","","","",""],
    #     ["N","Black Knight Helm","head","","","","","","","","","","","","","",""],
    #     ["N","Black Knight Leggings","legs","","","","","","","","","","","","","",""],
    #     ["Y","Black Leather Armor","chest",5.9,8.7,7.8,9.4,6.9,8.6,7.8,10.4,8.2,46,65,55,46,5.2],
    #     ["Y","Black Leather Boots","legs",3.6,4.8,4.3,5.2,3.8,4.8,4.3,5.6,4.3,25,37,31,25,2.3],
    #     ["Y","Black Leather Gloves","hands",2.3,2.2,2,2.3,1.8,2.1,1.9,2.9,2.3,14,20,18,14,1.7],
    #     ["N","Black Leggings","legs","","","","","","","","","","","","","",""],
    #     ["Y","Black Witch Garb","chest",3.6,3.7,4.2,4.2,2.7,13.8,13.5,13.5,14.1,15,39,28,69,0.5],
    #     ["Y","Black Witch Hat","head",1.8,1.9,1.1,1.1,1.4,4.9,4.8,4.8,5,6,17,12,30,0.2],
    #     ["Y","Black Witch Trousers","legs",2.6,2.9,2.4,2.4,2.1,8.1,7.9,7.9,8.3,11,25,19,45,0.5],
    #     ["Y","Black Witch Veil","head",1.4,0.8,0.9,0.8,0.8,4.7,4.5,4.6,4.9,4,14,8,44,0],
    #     ["Y","Black Witch Wrappings","hands",1.2,0.5,1.2,1.2,0.7,3.4,3.3,3.3,3.5,5,13,9,23,0.1],
    #     ["Y","Blindfold Mask","head",1.1,1,0.9,0.8,0.9,3.6,2.8,4.4,-30,6,15,10,32,0],
    #     ["Y","Brass Armor","chest",10.5,13.3,10.7,13.3,11.8,13.6,10,6.9,10,60,33,36,27,9.9],
    #     ["Y","Brass Gauntlets","hands",3.4,3.5,2.8,3.5,3.1,3.6,2.7,1.9,2.7,19,10,11,8,2.4],
    #     ["Y","Brass Helm","head",4.7,4.7,3.7,4.7,4.1,4.7,3.4,2.3,3.4,27,15,17,13,3.1],
    #     ["Y","Brass Leggings","legs",6.7,7.8,6.3,7.8,6.9,8.1,6,4.2,6,39,22,24,18,6.3],
    #     ["Y","Brigand Armor","chest",4.8,9,8.1,9,9,5.9,7.3,8.2,5.9,37,59,57,37,2.9],
    #     ["Y","Brigand Gauntlets","hands",2.4,2.5,2.3,2.5,2.5,1.9,2.1,2.3,1.9,14,20,19,14,1],
    #     ["Y","Brigand Hood","head",2.7,3.5,3.2,3.5,3.5,2.7,3,3.3,2.7,19,29,28,19,1.4],
    #     ["Y","Brigand Trousers","legs",5,6,5.6,6,6,4.8,5.2,5.6,4.8,32,43,41,32,3.1],
    #     ["N","Catarina Armor","chest",17.1,18.3,13.9,18.5,17.5,11.9,12.9,15,13.3,74,57,57,33,23.8],
    #     ["N","Catarina Gauntlets","hands",5.7,4.8,3.7,4.9,4.8,3.1,3.3,3.9,3.4,29,22,22,14,5.3],
    #     ["N","Catarina Helm","head",7.3,6.8,5,6.8,6.3,4.5,4.9,5.7,5.1,34,27,27,17,8.6],
    #     ["N","Catarina Leggings","legs",10.6,10.6,8.3,10.9,10.6,6.3,6.7,8,6.9,47,35,35,21,14.8],
    #     ["Y","Cathedral Knight Armor","chest",17,17.8,18,15.6,18.8,14.8,13.6,14.4,12.7,66,62,62,47,23.6],
    #     ["N","Cathedral Knight Gauntlets","hands","","","","","","","","","","","","","",""],
    #     ["N","Cathedral Knight Helm","head","","","","","","","","","","","","","",""],
    #     ["Y","Cathedral Knight Leggings","legs",9.7,8.9,9.1,7.7,10.3,7.6,7,7.4,6.2,40,38,38,26,13],
    #     ["Y","Chain Armor","chest",8.8,11.1,6.8,13.2,10,8.3,6.6,4.2,9,63,33,33,22,6.8],
    #     ["Y","Chain Helm","head",3.9,3.8,2.2,4.6,3.4,2.6,2,1.1,2.9,28,15,15,11,1.9],
    #     ["Y","Chain Leggings","legs",5,6,5.6,6,6,4.8,5.2,5.6,4.8,32,43,41,32,3.2],
    #     ["Y","Clandestine Coat","chest",3,4.2,4.2,4.2,3.6,13.5,12.9,13.1,13.5,25,36,30,65,0.9],
    #     ["Y","Cleric Blue Robe","chest",6,4.9,7.3,6,4.2,12.8,12.3,13.2,13.2,45,61,4.9,6.8,1.4],
    #     ["Y","Cleric Gloves","hands",1.5,1.1,0.7,0.7,0.9,2.8,2.1,2.8,2.9,17,21,11,25,0.5],
    #     ["Y","Cleric Hat","head",1.4,1.9,1.4,1.4,1.6,5,3.8,4.9,5,18,25,12,29,0.5],
    #     ["Y","Cleric Trousers","legs",2.1,2.4,1.6,1.6,1.9,6.9,5,6.8,7,26,36,18,42,0.7],
    #     ["Y","Conjurator Boots","legs",2.6,2.3,2.3,2.3,2,7,8.1,7.9,7.3,19,39,27,38,1],
    #     ["Y","Conjurator Hood","head",1.8,1.5,1.5,1.5,1.3,4.4,5.1,5,4.6,13,27,19,26,0.6],
    #     ["Y","Conjurator Manchettes","hands",1.4,0.8,0.8,0.8,1.6,2.7,3.2,3.1,2.9,10,21,15,20,0.1],
    #     ["Y","Conjurator Robe","chest",4.2,3.8,3.8,3.8,3.2,11.7,13.6,13.3,12.2,31,63,44,61,1.1],
    #     ["Y","Cornyx's Garb","chest",4.1,4.1,4.1,4.8,3.5,13,13.7,12.8,13.2,22,40,30,88,4.9],
    #     ["Y","Cornyx's Skirt","legs",2,1.7,1.7,1.7,1.5,6.9,7.1,6.6,7,15,24,18,48,1.4],
    #     ["Y","Cornyx's Wrap","hands",1.3,0.9,0.9,1.1,0.8,3.6,3.4,3.3,3.4,7,13,11,26,0.8],
    #     ["Y","Court Sorcerer Gloves","hands",1,1,1.1,1,1.1,3.2,3,3.6,3.5,7,12,7,21,0.5],
    #     ["N","Court Sorcerer Hood","head","","","","","","","","","","","","","",""],
    #     ["N","Court Sorcerer Robe","chest","","","","","","","","","","","","","",""],
    #     ["Y","Court Sorcerer Trousers","legs",2.2,2.2,2.5,2.2,2.5,7.4,6.8,8,7.9,19,29,21,47,1.1],
    #     ["N","Creighton's Steel Mask","head","","","","","","","","","","","","","",""],
    #     ["Y","Crown of Dusk","head",1,0.8,0,0.2,0,-30,4,3,3.2,5,12,8,33,0],
    #     ["Y","Dancer's Armor","chest",7.3,10.5,7.4,12.2,11.7,8.3,6,9.1,9.8,41,33,55,22,6.7],
    #     ["Y","Dancer's Crown","head",2.8,3.3,1.8,4.4,4,2.7,2.2,2.5,2.8,14,10,23,7,1.5],
    #     ["Y","Dancer's Gauntlets","hands",2.4,2.8,2,3.2,3.1,2.2,1.7,2.4,2.6,15,12,20,9,1.7],
    #     ["Y","Dancer's Leggings","legs",4.4,6,4.2,7,6.7,4.7,3.4,4.7,5.1,22,17,31,10,3.6],
    #     ["Y","Dark Armor","chest",9.1,12.6,9.2,11.6,11.1,9.2,7.7,4.4,9.2,38,30,33,14,8.8],
    #     ["N","Dark Gauntlets","hands","","","","","","","","","","","","","",""],
    #     ["N","Dark Leggings","legs","","","","","","","","","","","","","",""],
    #     ["N","Dark Mask","head","","","","","","","","","","","","","",""],
    #     ["Y","Deacon Robe","chest",3.5,3.6,3.6,4.4,3.6,12.5,12.7,11.6,13.2,19,40,33,79,1],
    #     ["Y","Deacon Skirt","legs",2.3,1.1,1.7,2.1,1.7,6.8,6.9,6.3,7.2,13,26,22,50,1.1],
    #     ["Y","Desert Pyromancer  Skirt","legs",2.4,3.5,1.6,1.9,1.9,6.2,7.3,7.1,7.4,13,43,18,44,1.2],
    #     ["Y","Desert Pyromancer Garb","chest",2.3,3.7,2.9,2.4,2.4,10.5,11.1,12.7,11.7,16,64,25,66,0.9],
    #     ["Y","Desert Pyromancer Gloves","hands",1.1,1.1,1,1,0.9,3,2.8,3.4,2.9,9,25,11,25,0.3],
    #     ["Y","Desert Pyromancer Hood","head",1.2,1.9,0.8,1.2,1.2,3.6,4.5,4.2,4.7,10,30,13,31,0.3],
    #     ["Y","Deserter Armor","chest",8.6,11.7,9.3,11.7,11.1,6.6,7.5,4.2,10.4,50,33,36,22,8.6],
    #     ["Y","Deserter Trousers","legs",3.1,4.2,4.7,4.2,4.2,5,5,5.9,7,21,34,28,25,2.1],
    #     ["N","Dragonscale Armor","chest","","","","","","","","","","","","","",""],
    #     ["N","Dragonscale Waistcloth","","","","","","","","","","","","","","",""],
    #     ["Y","Dragonslayer Armor","chest",13.8,15.6,12.7,17,17,12.9,13.8,13.8,10.7,64,40,43,36,19.1],
    #     ["Y","Dragonslayer Gauntlets","hands",4.1,3.7,3,4.1,4.1,3.3,3.7,3.7,2.6,21,13,14,12,4.2],
    #     ["Y","Dragonslayer Helm","head",5.6,5,4,5.5,5.5,3.7,4.1,4.1,2.9,28,18,18,16,5.5],
    #     ["Y","Dragonslayer Leggings","legs",8.1,8.1,6.5,8.9,8.9,6.3,6.9,6.9,4.9,38,23,24,21,10.3],
    #     ["N","Drakeblood Armor","chest","","","","","","","","","","","","","",""],
    #     ["N","Drakeblood Gauntlets","hands","","","","","","","","","","","","","",""],
    #     ["N","Drakeblood Helm","head","","","","","","","","","","","","","",""],
    #     ["N","Drakeblood Leggings","legs","","","","","","","","","","","","","",""],
    #     ["Y","Drang Armor","chest",5.1,9.3,8.6,7.7,5.5,9.7,7.5,10.9,9.7,42,62,66,45,0.8],
    #     ["Y","Drang Gauntlets","hands",1.7,2.1,1.9,1.7,1.1,2,1.4,2.3,2,13,19,21,41,0.3],
    #     ["Y","Drang Shoes","legs",4.2,4.7,5.2,5.2,5.2,6.3,5,7,6.3,24,37,39,26,1.3],
    #     ["Y","Eastern Armor","chest",10.8,12.6,9.3,16.9,11.5,7.6,11.5,5.3,6.7,87,37,43,25,12.6],
    #     ["Y","Eastern Gauntlets","hands",2.9,2.8,2.2,3.8,2.5,1.8,3.1,1.5,1.8,25,9,11,7,2.6],
    #     ["Y","Eastern Helm","head",3.9,4.1,2.5,5.6,3.7,1.6,3.4,1.2,1.6,32,12,15,9,3.3],
    #     ["Y","Eastern Leggings","legs",5,5.1,4.7,7.5,4.6,5.4,7.2,6.6,5.4,47,40,41,39,4.3],
    #     ["Y","Elite Knight Armor","chest",8.9,12.1,9.2,12.1,11.1,9.2,10.6,6.8,8.5,46,32,33,18,8.8],
    #     ["Y","Elite Knight Gauntlets","hands",3.4,3.4,2.8,3.5,3.4,2.6,3,2,2.4,18,13,14,9,2.3],
    #     ["Y","Elite Knight Helm","head",5.2,4.9,3.9,4.9,4.5,3.9,4.2,3.3,3.7,24,15,15,9,4.7],
    #     ["Y","Elite Knight Leggings","legs",6.8,8.3,6.8,8.3,7.4,6.8,7.4,5.7,6.5,38,25,26,16,7],
    #     ["Y","Embraced Armor of Favor","chest",12.8,13.2,12.1,13.2,13.7,11.1,11.6,8.5,11.6,63,46,51,35,18.3],
    #     ["N","Evangelist Gloves","hands","","","","","","","","","","","","","",""],
    #     ["Y","Evangelist Hat","head",3.5,3.5,4.3,3.2,2.9,4.1,4.6,4.6,2.9,15,27,24,25,2.1],
    #     ["N","Evangelist Robe","chest","","","","","","","","","","","","","",""],
    #     ["Y","Evangelist Trousers","legs",4.6,5.5,6.9,5.1,4.6,6.6,7.3,7.3,4.6,22,39,35,37,3.6],
    #     ["Y","Executioner Armor","chest",16,18.3,12.4,17.5,18.3,13.5,14.5,11.4,13.3,65,58,61,42,21.4],
    #     ["Y","Executioner Gauntlets","hands",5.3,4.6,3.1,4.6,4.2,3.3,3.6,2.8,3.3,24,22,23,17,4.9],
    #     ["Y","Executioner Helm","head",6.8,6.7,4.4,6.7,6.1,4.9,5,4,4.7,31,28,29,21,7.5],
    #     ["Y","Executioner Leggings","legs",9.9,10.6,7.1,10.1,10.6,7.9,8.1,6.5,7.6,39,35,37,25,13.2],
    #     ["Y","Exile Armor","chest",17.3,18.2,17.6,16.6,15.9,12.4,14.3,12.6,13.1,97,74,62,50,23.1],
    #     ["Y","Exile Gauntlets","hands",5.5,4.8,4.4,4.1,4,3.3,3.7,3.3,3.4,33,25,22,17,5],
    #     ["Y","Exile Leggings","legs",10.4,10.8,10.4,9.8,9.3,7.4,8.5,7.5,7.8,57,43,35,28,14.8],
    #     ["Y","Exile Mask","head",7.1,6.5,6.3,7,5.9,4.4,5,3.6,4.5,39,29,24,19,8.7],
    #     ["Y","Fallen Knight Armor","chest",9.2,12.7,12.2,10,11.1,8.3,12.5,10.4,9,43,28,39,22,8.7],
    #     ["Y","Fallen Knight Gauntlets","hands",3.1,3.3,3.2,2.6,2.9,2.4,3.7,3,2.2,11,5,10,4,2.1],
    #     ["Y","Fallen Knight Helm","head",4.6,4.6,4.5,3.6,4,3.1,4.2,3.5,2.6,21,14,19,12,3.5],
    #     ["Y","Fallen Knight Trousers","legs",5.3,7.3,7,5.6,6.4,5,7.8,6.3,5.5,22,10,20,9,4.9],
    #     ["Y","Faraam Armor","chest",12.7,13.9,13.8,13.9,12.8,10.7,12.2,11.2,11.2,59,40,60,27,17.3],
    #     ["Y","Faraam Boots","legs",6.7,7.2,7.5,7.2,6.6,5.6,6.4,5.6,5.6,34,24,37,15,8.2],
    #     ["Y","Faraam Gauntlets","hands",4.7,3.6,3.7,3.6,3.3,3,3.3,3.1,3.1,21,16,22,11,4.5],
    #     ["Y","Faraam Helm","head",6.3,5.6,5.5,5.6,5.1,4.2,4.4,4.3,4.3,27,22,28,14,7],
    #     ["Y","Fire Keeper Robe","chest",5.1,5.7,7.1,7.1,8,7.1,11.9,8.8,12.4,36,54,49,42,1.9],
    #     ["Y","Fire Keeper Skirt","legs",2.1,1.9,2.6,1.9,1.9,7.1,7.6,6.8,8.3,13,24,18,40,0.5],
    #     ["Y","Fire Kepper Gloves","hands",1.3,0.8,0.9,1.1,0.7,3.1,3.3,2.9,3.6,9,15,11,23,0.2],
    #     ["Y","Fire Witch Armor","chest",10.9,13.2,8.5,10.6,10.6,12.1,12.3,8.5,11.1,47,34,36,27,9.8],
    #     ["N","Fire Witch Gauntlets","hands","","","","","","","","","","","","","",""],
    #     ["N","Fire Witch Helm","head","","","","","","","","","","","","","",""],
    #     ["N","Fire Witch Leggings","legs","","","","","","","","","","","","","",""],
    #     ["N","Firelink Armor","chest","","","","","","","","","","","","","",""],
    #     ["N","Firelink Gauntlets","hands","","","","","","","","","","","","","",""],
    #     ["N","Firelink Helm","helm","","","","","","","","","","","","","",""],
    #     ["N","Firelink Leggings","legs","","","","","","","","","","","","","",""],
    #     ["Y","Follower Armor","chest",5.6,7.1,10.2,8,9.5,10.2,8.8,10.9,8.8,37,59,65,48,2.7],
    #     ["Y","Follower Boots","legs",3.3,4,5.8,4.5,5.4,5.8,5,6.2,5,23,36,40,30,1.6],
    #     ["Y","Follower Gloves","hands",1.7,1.7,2.5,1.9,2.3,2.5,2.1,2.7,2.1,12,20,22,16,0.6],
    #     ["N","Follower Helm","head","","","","","","","","","","","","","",""],
    #     ["Y","Gauntlets of Favor","hands",4,3.5,3.2,3.5,3.5,3,3.1,2.2,3.1,19,13,14,10,4],
    #     ["Y","Gauntlets of Thorns","hands",2.8,2.8,2.1,3.6,2.9,1.9,1.9,1.9,1.6,21,8,10,7,1.8],
    #     ["N","Golden Bracelets","hands","","","","","","","","","","","","","",""],
    #     ["N","Golden Crown","head","","","","","","","","","","","","","",""],
    #     ["Y","Grave Warden Hood","head",1.5,1.5,1.2,2.9,1.8,3.3,3.3,4,1.4,30,28,8,25,0.5],
    #     ["N","Grave Warden Robe","chest","","","","","","","","","","","","","",""],
    #     ["N","Grave Warden Skirt","legs","","","","","","","","","","","","","",""],
    #     ["Y","Grave Warden Wrap","hands",1.2,0.8,0.5,1.8,1,1.9,1.9,2.5,0.6,20,18,3,16,0],
    #     ["Y","Gundyr's Armor","chest",17.2,16.9,14.1,19.7,16.2,13.2,13.7,13,13.2,86,62,62,62,24.6],
    #     ["Y","Gundyr's Gauntlets","hands",5.8,4.4,3.6,5.1,4.1,3.4,3.5,3.4,3.4,29,21,21,21,6],
    #     ["Y","Gundyr's Helm","head",7.2,5.9,4.8,7,5.6,4.4,4.6,4.4,4.4,35,25,25,25,8.3],
    #     ["Y","Gundyr's Leggings","legs",9.8,9,7.5,10.7,8.6,7.1,7.4,7,7.1,48,36,36,36,14.4],
    #     ["N","Harald Legion Armor","chest","","","","","","","","","","","","","",""],
    #     ["N","Harald Legion Gauntlets","hands","","","","","","","","","","","","","",""],
    #     ["N","Harald Legion Leggings","legs","","","","","","","","","","","","","",""],
    #     ["N","Hard Leather Armor","chest","","","","","","","","","","","","","",""],
    #     ["Y","Hard Leather Boots","legs",4.7,5.1,5.5,5.1,5.1,4.1,4.6,2.8,4.1,26,38,25,29,2.9],
    #     ["Y","Hard Leather Gauntlets","hands",2,2.1,2.2,2.1,2.1,1.5,1.7,1,1.5,16,22,21,17,0.8],
    #     ["N","Havel's Armor","Helm","","","","","","","","","","","","","",""],
    #     ["N","Havel's Gauntlets","hands","","","","","","","","","","","","","",""],
    #     ["N","Havel's Helm","head","","","","","","","","","","","","","",""],
    #     ["N","Havel's Leggings","legs","","","","","","","","","","","","","",""],
    #     ["Y","Helm of Favor","head",5.7,4.5,4.1,4.5,4.7,3.7,3.9,2.8,3.9,30,23,25,18,6],
    #     ["Y","Helm of Thorns","head",3.8,3.7,2.6,4.9,3.9,2.3,2.3,2.3,1.8,30,13,15,11,2.2],
    #     ["Y","Herald Armor","chest",8.6,12.4,8,12.9,11.9,9.2,8.4,3.8,8.4,52,30,36,19,9.5],
    #     ["Y","Herald Gloves","hands",2.9,3.1,2.1,3.1,2.9,1.7,2.1,0.7,1.5,16,12,12,8,1.6],
    #     ["Y","Herald Helm","head",3.7,4.2,2.5,4.4,4,3.5,3.2,1.6,3.2,24,14,17,10,3.4],
    #     ["Y","Herald Trousers","legs",5.3,7.1,6.5,6.8,6.5,4.6,4.6,1.8,4.1,26,20,20,11,5.1],
    #     ["Y","Hood of Prayer","head",1.3,0.8,0.8,0.8,0.8,3.9,3.5,3.7,4.4,8,28,13,40,0.2],
    #     ["Y","Iron Bracelets","hands",2.9,2.4,1.6,2.4,2.4,2.2,2.2,2,2.2,11,9,10,9,2.2],
    #     ["Y","Iron Dragonslayer Armor","chest",14.8,15.7,11.6,16.5,17.2,9.9,11.6,13.8,13.2,81,54,60,29,20.1],
    #     ["Y","Iron Dragonslayer Gauntlets","hands",5.1,3.7,2.6,3.9,4.1,2.2,2.6,3.2,3,26,17,19,8,4.8],
    #     ["Y","Iron Dragonslayer Helm","head",6.3,6.1,4.5,6.4,6.6,3.9,4.5,5.3,5.1,33,22,24,11,7],
    #     ["Y","Iron Dragonslayer Leggings","legs",9.5,9.5,7.1,9.9,10.4,6.1,7.1,8.4,8,48,32,35,16,12.5],
    #     ["Y","Iron Helm","head",5.8,4.9,5.5,4.7,4.7,4.8,4.8,4.7,4.8,17,17,17,15,6.6],
    #     ["N","Iron Leggings","legs","","","","","","","","","","","","","",""],
    #     ["Y","Jailer Gloves","hands",1.6,1.7,0.7,1.7,1.1,2.7,2.6,2.2,2.7,10,15,13,32,0.3],
    #     ["Y","Jailer Robe","chest",4.8,8.2,4.2,8.22,6,13.5,13.1,11.4,13.5,24,37,32,87,2],
    #     ["Y","Jailer Trousers","legs",2.8,4.2,1.9,4.2,3,6.9,6.7,5.7,6.9,12,22,19,53,1],
    #     ["Y","Karla's Coat","chest",3.6,3.9,3.9,3.9,2.9,14.1,13.3,13.8,13.8,20,35,31,53,1.3],
    #     ["Y","Karla's Gloves","hands",1.2,0.8,0.9,0.9,0.9,3.5,3.3,3.4,3.4,7,12,10,18,0.3],
    #     ["Y","Karla's Pointed Hat","head",1.7,1.6,1.4,1.4,1.2,5,4.7,4.9,4.9,9,15,13,23,0.4],
    #     ["Y","Karla's Trousers","legs",2.6,2.6,2.6,2.6,1.9,8.3,7.7,8.1,8.1,14,23,21,36,1.1],
    #     ["Y","Knight Armor","chest",10.6,13.8,11.7,13.8,13.2,8.4,11,6.7,8.4,45,36,28,24,12.6],
    #     ["Y","Knight Gauntlets","hands",3.5,3.7,3.2,3.7,3.6,2.5,2.9,2.1,2.5,14,12,11,7,2.5],
    #     ["Y","Knight Helm","head",5.1,5.3,4.2,5.3,4.8,3,4,2.7,3.2,22,17,18,12,4.6],
    #     ["Y","Knight Leggings","legs",6.6,8.3,7,8.3,7.9,5.4,6.2,4.4,5.4,26,23,22,13,7.6],
    #     ["N","Lapp's Armor","chest","","","","","","","","","","","","","",""],
    #     ["N","Lapp's Gauntlets","hands","","","","","","","","","","","","","",""],
    #     ["N","Lapp's Helm","head","","","","","","","","","","","","","",""],
    #     ["N","Lapp's Leggings","legs","","","","","","","","","","","","","",""],
    #     ["Y","Leather Armor","chest",5.4,6.8,6.8,6.8,6.8,8.3,9,9.7,8.3,36,58,58,43,3.3],
    #     ["Y","Leather Boots","legs",3.3,4.2,4.2,4.2,4.2,5.5,5.9,6.3,5.5,24,37,37,28,2.8],
    #     ["Y","Leather Gauntlets","hands",2.5,2.2,2,2.9,2,3.2,2.6,3.2,3.3,18,18,16,16,1.1],
    #     ["Y","Leather Gloves","hands",1.5,1.4,1.4,1.4,1.4,1.9,2.1,2.3,1.9,11,18,17,13,0.1],
    #     ["Y","Leggings of Favor","legs",7.5,7.7,7.1,7.7,7.7,6.5,6.8,4.8,6.8,41,29,32,24,10.3],
    #     ["Y","Leggings of Thorns","legs",5.4,6,4.3,7.9,6.3,3.8,3.8,3.8,3,42,18,22,15,4.2],
    #     ["N","Leohard's Garb","chest","","","","","","","","","","","","","",""],
    #     ["N","Leohard's Gauntlets","hands","","","","","","","","","","","","","",""],
    #     ["N","Leohard's Trousers","legs","","","","","","","","","","","","","",""],
    #     ["Y","Loincloth","legs",1.1,1.9,1.9,1.9,1.6,5.8,4.9,5.4,5.8,20,35,10,43,0.4],
    #     ["Y","Loincloth (DLC)","legs",2,2.4,2.1,1.9,1.7,7.7,6.7,7.3,7.4,18,29,15,37,0.6],
    #     ["Y","Lorian's Armor","chest",14.7,16.6,10.1,16.6,15.9,13.3,11.2,11.2,13.6,67,41,61,50,18.4],
    #     ["Y","Lorian's Gauntlets","hands",3,3.2,2.3,3,2.9,2.9,1.5,1.5,2.9,16,7,11,10,1.8],
    #     ["Y","Lorian's Helm","head",5.5,5.4,3.5,5.4,3.8,4.8,3.6,3.6,4.8,24,12,17,15,4.6],
    #     ["Y","Lorian's Leggings","legs",6.9,8.2,5.9,8.2,7.5,7.4,5.2,5.2,7.5,38,21,27,25,6.3],
    #     ["Y","Lothric Knight Armor","chest",15,16,12.4,18.4,15.3,12.6,13.8,10.2,11.9,75,43,58,34,20.8],
    #     ["Y","Lothric Knight Gauntlets","hands",4.8,3.8,2.9,4.6,3.6,3.1,3.3,2.3,2.8,25,15,19,12,4.5],
    #     ["Y","Lothric Knight Helm","head",5.8,5.2,4,6.3,5,4.2,4.7,3.1,3.8,31,18,22,14,6.3],
    #     ["Y","Lothric Knight Leggings","legs",8.9,8.8,6.8,10.6,8.4,7.1,7.7,5.4,6.5,44,24,32,20,12.1],
    #     ["N","Lucatiel's Mask","head","","","","","","","","","","","","","",""],
    #     ["Y","Maiden Gloves","hands",1.3,0.9,1.1,0.9,0.8,3.5,3.3,3.4,3.5,8,14,11,26,0.3],
    #     ["Y","Maiden Hood","head",1.4,1,1.2,1,0.9,4.6,4.3,4.4,4.6,8,16,12,31,0],
    #     ["Y","Maiden Robe","chest",3.5,3.5,4.1,3.5,3.1,13.7,12.8,13,13.7,19,38,30,75,0.7],
    #     ["Y","Maiden Skirt","legs",2.3,1.7,2,1.7,1.5,7.5,7,7.1,7.5,13,25,20,48,0.1],
    #     ["Y","Master's Attire","chest",2,2.7,2.7,2.7,2.1,9.1,7.6,8.4,9.1,37,63,22,74,0],
    #     ["Y","Master's Gloves","hands",0.3,0.4,0.4,0.4,0.3,1.9,1.5,1.7,1.9,14,22,9,26,0],
    #     ["N","Millwood Knight Armor","chest","","","","","","","","","","","","","",""],
    #     ["N","Millwood Knight Gauntlets","hands","","","","","","","","","","","","","",""],
    #     ["N","Millwood Knight Helm","head","","","","","","","","","","","","","",""],
    #     ["N","Millwood Knight Leggings","legs","","","","","","","","","","","","","",""],
    #     ["N","Mirrah Chain Gloves","hands","","","","","","","","","","","","","",""],
    #     ["N","Mirrah Chain Leggings","legs","","","","","","","","","","","","","",""],
    #     ["N","Mirrah Chain Mail","chest","","","","","","","","","","","","","",""],
    #     ["Y","Mirrah Gloves","hands",2.1,2.1,2.1,1.5,1.5,1.9,1.7,2.5,1.7,17,21,17,19,0.2],
    #     ["Y","Mirrah Trousers","legs",3.5,5.1,5.1,3.7,3.7,5,4.6,6.4,4.6,29,37,29,32,1.2],
    #     ["Y","Mirrah Vest","chest",7,10.5,10.5,8.2,8.2,10.6,9.9,12.8,9.9,54,64,54,60,4.2],
    #     ["Y","Morne's Armor","chest",15.5,17.4,14.3,17.4,16.6,13.9,13.6,13.3,9.3,67,39,60,37,20.6],
    #     ["Y","Morne's Gauntlets","hands",4.8,3.8,3.2,3.8,3.7,3.1,3.1,3,1.7,21,12,18,11,4.1],
    #     ["Y","Morne's Helm","head",7.7,6.8,5.8,6.8,6.8,5.5,5.4,5.2,4,31,17,27,17,8],
    #     ["Y","Morne's Leggings","legs",9.3,9.8,8.1,9.8,9.3,8.1,7.9,7.8,5.2,40,23,36,22,12.8],
    #     ["Y","Nameless Knight Armor","chest",9.3,11.7,10,12.7,11.1,7.6,11,8.4,7.6,61,38,35,21,8.7],
    #     ["Y","Nameless Knight Helm","head",3.8,3.8,3.1,4.2,3.6,2.1,3.4,2.4,2.1,26,15,14,10,2.3],
    #     ["Y","Nameless Knight Leggings","legs",5.6,7,6,7.6,6.7,4.9,6.9,5.4,4.9,36,22,20,11,5.9],
    #     ["Y","Nameless Knight Gauntlets","hands",2.8,3,2.6,3.3,2.9,2.1,3.1,2.3,2.1,19,10,9,6,2.2],
    #     ["Y","Northern Armor","chest",10.6,12.6,12.6,12.1,11.6,7.7,11.1,4.4,6.8,60,38,64,27,10.7],
    #     ["Y","Northern Gloves","hands",2.3,2.2,2.9,2.2,2.2,0.8,2.5,2,0.8,15,19,24,19,0.7],
    #     ["Y","Northern Helm","head",4.8,4.8,4.7,4.6,4.4,2.9,4.2,1.8,2.6,24,15,26,10,4],
    #     ["Y","Northern Trousers","legs",4.3,6.1,7.6,6.1,6.1,2.9,6.8,5.7,2.9,24,32,42,32,3.1],
    #     ["Y","Old Sage's Blindfold","head",1,0.8,0.8,0.8,0.8,4,4.4,3.8,4.2,10,14,11,28,0],
    #     ["Y","Old Sorcerer Boots","legs",2.3,2.1,1.8,1.8,2.1,7.3,6.7,7,7.3,18,23,20,42,1.1],
    #     ["Y","Old Sorcerer Coat","chest",3.7,3.6,3.1,3.1,3.6,12.7,11.6,12.1,12.7,32,40,35,70,1],
    #     ["Y","Old Sorcerer Gauntlets","hands",1.3,0.9,0.8,0.8,0.9,3.2,2.9,3.1,3.2,11,14,12,24,0],
    #     ["Y","Old Sorcerer Hat","head",1.2,1.5,1.3,1.3,1.5,4.9,4.3,4.5,4.9,11,15,12,28,0.6],
    #     ["Y","Ordained Dress","chest",4.4,5.7,3.9,8,2.9,14.5,11.4,12.6,12.6,20,53,48,69,2.7],
    #     ["Y","Ordained Hood","head",1.9,2,1.4,2.8,1,5.2,4,4.6,4.5,9,19,21,32,0.4],
    #     ["Y","Ordained Trousers","legs",2.8,3.2,2.2,4.5,1.6,8.3,6.5,7.2,7.2,12,33,30,43,1.6],
    #     ["Y","Outrider Knight Armor","chest",12,14.4,10.6,13.8,13.8,9.7,11.4,9.7,7.5,50,43,91,22,12.5],
    #     ["Y","Outrider Knight Gauntlets","hands",2.9,3.4,2.4,3.3,3.4,2.4,3,2.4,1.6,12,10,25,4,2.7],
    #     ["Y","Outrider Knight Helm","head",3.8,4.4,2.8,4.2,4.4,2.3,3.1,2.3,1.1,17,12,32,5,2.8],
    #     ["Y","Outrider Knight Leggings","legs",6.8,7.9,6,7.6,7.9,5.9,6.7,5.5,3.7,32,24,55,14,7.2],
    #     ["N","Painting Guardian Gloves","hands","","","","","","","","","","","","","",""],
    #     ["N","Painting Guardian Gown","chest","","","","","","","","","","","","","",""],
    #     ["N","Painting Guardian Hood","head","","","","","","","","","","","","","",""],
    #     ["N","Painting Guardian Waistcloth","legs","","","","","","","","","","","","","",""],
    #     ["N","Pale Shadow Gloves","hands","","","","","","","","","","","","","",""],
    #     ["N","Pale Shadow Robe","chest","","","","","","","","","","","","","",""],
    #     ["N","Pale Shadow Trousers","legs","","","","","","","","","","","","","",""],
    #     ["N","Pontiff Knight Armor","chest","","","","","","","","","","","","","",""],
    #     ["N","Pontiff Knight Crown","head","","","","","","","","","","","","","",""],
    #     ["N","Pontiff Knight Gauntlets","hands","","","","","","","","","","","","","",""],
    #     ["Y","Pontiff Knight Leggings","legs",4.3,5.1,3,5.1,4.7,7.6,5.5,7,7.8,23,21,48,16,2.8],
    #     ["Y","Pyromancer Crown","head",1.1,0.7,0.7,0.7,0.7,4.2,4.8,3.6,4,9,14,15,21,0],
    #     ["Y","Pyromancer Garb","chest",4.2,4.4,3.7,3.7,3.7,12,13.1,12,14.1,22,50,43,74,0.9],
    #     ["Y","Pyromancer Trousers","legs",2.6,2.7,2.2,2.2,2.2,7.1,8.1,7.5,7.8,14,43,26,47,1.2],
    #     ["Y","Pyromancer Wrap","hands",1.5,1.4,1.4,1.2,1.2,3.5,3.6,3.5,3.6,6,14,11,22,0.8],
    #     ["Y","Ragged Mask","head",0.7,1.1,1.1,1.1,1.1,3.8,3.2,3.5,3.8,13,25,7,29,0.4],
    #     ["N","Ringed Knight Armor","chest","","","","","","","","","","","","","",""],
    #     ["N","Ringed Knight Gauntlets","hands","","","","","","","","","","","","","",""],
    #     ["N","Ringed Knight Hood","head","","","","","","","","","","","","","",""],
    #     ["N","Ringed Knight Leggings","legs","","","","","","","","","","","","","",""],
    #     ["Y","Robe of Prayer","chest",3.4,3.6,3.2,3.6,3.2,13.5,13.1,13.1,14.6,13,61,24,89,1.3],
    #     ["Y","Ruin Armor","chest",12.9,13.7,13.1,16.3,13.7,12.8,11.6,12.2,10.5,62,60,35,32,17.4],
    #     ["Y","Ruin Gauntlets","hands",4.2,3.6,3.4,4.2,3.6,3.4,3,3.2,2.7,19,19,10,9,4.1],
    #     ["Y","Ruin Helm","head",5.4,4.6,4.3,5.5,4.6,4.2,3.8,4,3.4,25,24,13,12,6.1],
    #     ["Y","Ruin Leggings","legs",8,7.4,7.1,8.9,7.4,6.9,6.2,6.5,5.5,37,35,20,18,10.7],
    #     ["Y","Sage's Big Hat","head",1.9,2.5,2,2,1.6,5.2,3.6,4.6,4.6,15,21,13,26,0.4],
    #     ["N","Scholar's Robe","chest","","","","","","","","","","","","","",""],
    #     ["Y","Sellsword Armor","chest",10.6,12.2,12.7,12.2,12.2,11.1,9.1,7.3,10.5,59,39,35,26,11.5],
    #     ["Y","Sellsword Gauntlet","hands",1.4,1.4,1.4,0.8,1.6,0.6,0.8,0.8,0.6,15,10,11,8,0.3],
    #     ["Y","Sellsword Helm","head",3.1,3.1,3.3,3.6,3.9,2.7,2.7,2.7,2.7,27,17,16,12,3.5],
    #     ["Y","Sellsword Trousers","legs",3.6,4.7,5.1,4.2,4.7,2.3,3.7,2.3,2.3,28,17,19,9,3.4],
    #     ["Y","Shadow Garb","chest",3.7,6.8,5.5,4.4,6.8,4.2,5.2,6.6,4.2,27,40,38,34,2.2],
    #     ["Y","Shadow Gauntlets","hands",1.3,1.6,1.4,1.6,1.6,1.6,1.9,2.2,1.6,10,15,14,13,1.1],
    #     ["Y","Shadow Leggings","legs",2.3,4.2,3.5,2.9,4.2,3.1,3.7,4.5,3.1,16,24,23,21,2.1],
    #     ["Y","Shadow Mask","head",1.5,2.5,2.2,1.7,2.5,2,1.5,2,1.1,15,18,20,18,0.4],
    #     ["N","Shira's Armor","chest","","","","","","","","","","","","","",""],
    #     ["N","Shira's Crown","head","","","","","","","","","","","","","",""],
    #     ["N","Shira's Gloves","hands","","","","","","","","","","","","","",""],
    #     ["N","Shira's Trousers","legs","","","","","","","","","","","","","",""],
    #     ["Y","Silver Knight Armor","chest",14.4,14.3,13.1,17,16.3,13.1,12.9,11.9,11.9,65,55,50,39,19.9],
    #     ["Y","Silver Knight Gauntlets","hands",4.6,3.6,3.1,4.3,4.1,3.1,2.9,2.7,2.7,22,18,17,14,3.8],
    #     ["N","Silver Knight Helm","head","","","","","","","","","","","","","",""],
    #     ["N","Silver Knight Leggings","legs","","","","","","","","","","","","","",""],
    #     ["N","Silver Mask","head","","","","","","","","","","","","","",""],
    #     ["Y","Skirt of Prayer","legs",2,1.6,1.9,1.6,1.4,6.9,6.7,6.7,7.6,14,39,18,58,0],
    #     ["Y","Slave Knight Armor","chest",8.7,13.5,10.9,13.5,10.2,5.7,8,5.7,7.1,61,34,35,20,9.1],
    #     ["Y","Slave Knight Gauntlets","hands",3,3.3,2.7,3.3,2.5,1.4,1.9,1.4,1.7,20,11,12,7,2],
    #     ["Y","Slave Knight Hood","head",3.8,4.8,3.8,4.8,3.6,2,2.8,2,2.5,26,14,15,9,3],
    #     ["Y","Slave Knight Leggings","legs",5.4,7.7,6.2,7.7,5.8,3.2,4.5,3.2,4,38,21,22,12,5.5],
    #     ["Y","Smough's Armor","chest",23,18.8,19.2,18.3,16,13.3,13,13.8,13.1,69,69,69,62,27.1],
    #     ["Y","Smough's Gauntlets","hands",9.4,4.8,4.9,4.6,4,3.3,3.2,3.4,3.5,20,20,20,18,6.9],
    #     ["Y","Smough's Helm","head",11.8,6.8,7,6.7,5.8,4.7,4.6,4.9,4.9,27,27,27,24,9.9],
    #     ["Y","Smough's Leggings","legs",15.8,10.9,11.1,10.6,9.2,7.6,7.4,7.9,7.6,43,43,43,39,17.3],
    #     ["N","Sneering Mask","head","","","","","","","","","","","","","",""],
    #     ["Y","Sorcerer Gloves","hands",1.3,1.3,1.3,1.3,1.1,3.5,3.3,3.4,3.5,8,13,10,23,0.6],
    #     ["Y","Sorcerer Hood","head",1.4,1,1,1,0.9,4.4,4,4.2,4.5,11,18,14,30,0.1],
    #     ["Y","Sorcerer Robe","chest",4.1,4.5,4.5,4.5,3.8,13.1,12.4,12.6,13.1,25,40,33,71,1.7],
    #     ["Y","Sorcerer Trousers","legs",3.2,3.4,4.1,3.4,3.4,6,7.5,7.6,6.7,22,26,31,28,1.3],
    #     ["Y","Standard Helm","head",3.7,3.7,3.7,3.7,3.7,3,3.3,2.4,3,22,29,27,24,2.5],
    #     ["Y","Steel Soldier Helm","head",4.6,4.4,3.8,4.4,4.2,2.9,2.9,2,3.7,27,18,18,13,3.2],
    #     ["N","Sunset Armor","chest","","","","","","","","","","","","","",""],
    #     ["N","Sunset Gauntlets","hands","","","","","","","","","","","","","",""],
    #     ["N","Sunset Helm","head","","","","","","","","","","","","","",""],
    #     ["N","Sunset Leggings","legs","","","","","","","","","","","","","",""],
    #     ["N","Sunless Armor","chest","","","","","","","","","","","","","",""],
    #     ["N","Sunless Gauntlets","hands","","","","","","","","","","","","","",""],
    #     ["N","Sunless Leggings","legs","","","","","","","","","","","","","",""],
    #     ["N","Sunless Veil","head","","","","","","","","","","","","","",""],
    #     ["N","Symbol of Avarice","head","","","","","","","","","","","","","",""],
    #     ["Y","Thief Mask","head",2.1,2.2,2.5,2.2,2.2,2.2,2.5,2.8,2.2,23,26,21,17,0.6],
    #     ["Y","Thrall Hood","head",1.5,2,1.4,1.6,1.2,4.2,3.8,4,4.2,17,28,19,38,0.8],
    #     ["Y","Undead Legion Armor","chest",7.6,10.4,10.4,5.9,9,12.6,13,13.2,10.4,39,51,65,42,4.9],
    #     ["Y","Undead Legion Gauntlets","hands",2.4,2.3,2.3,1.2,2,2.9,3,3,2.3,16,20,24,17,0.8],
    #     ["Y","Undead Legion Helm","head",4,3.8,4,2.6,3.5,4.7,4.8,4.9,4.2,19,24,29,21,2.3],
    #     ["Y","Undead Legion Leggings","legs",4.6,5.5,5.6,3.1,4.8,6.9,7.1,7.2,5.6,27,35,43,29,2.6],
    #     ["Y","Vilhelm's Armor","chest",10.8,12.9,10.2,12.9,14.6,9.5,12.4,5.7,11.4,44,35,37,31,13],
    #     ["Y","Vilhelm's Gauntlets","hands",3.7,3.2,2.5,3.2,3.6,2.3,3.1,1.4,2.8,15,12,12,10,2.9],
    #     ["Y","Vilhelm's Helm","head",4.8,4.6,3.6,4.6,5.2,3.4,4.4,2,4,19,15,16,13,4.4],
    #     ["Y","Vilhelm's Leggings","legs",6.8,7.4,5.8,7.4,8.4,5.4,7.1,3.2,6.5,27,22,23,19,8],
    #     ["Y","Violet Wrappings","hands",0.9,0.7,0.5,0.4,0.4,3,2.5,2.8,2.9,11,17,10,21,0.1],
    #     ["Y","White Preacher Head","head",3.5,2.6,4.2,4,4,4.6,3.8,4,2.9,17,28,24,29,2.6],
    #     ["Y","Winged Knight Armor","chest",18.2,17.5,14.4,17.5,16.8,14.3,14,15,16.7,73,60,68,45,24.9],
    #     ["Y","Winged Knight Gauntlets","hands",6,4.8,3.7,4.8,4.4,3.5,4,3.8,4.5,22,21,21,13,5.5],
    #     ["Y","Winged Knight Helm","head",7,6.7,5.4,6.7,6.3,4.3,5,4.7,5.7,31,29,29,18,6.9],
    #     ["Y","Winged Knight Leggings","legs",10.5,10.1,8.3,10.1,9.6,7.2,7.1,7.6,8.4,42,32,39,24,14.1],
    #     ["Y","Wolf Knight Armor","chest",9,11.8,10.1,11.8,11.8,7.8,11.2,5.5,10,46,32,33,33,8],
    #     ["Y","Wolf Knight Gauntlets","hands",3.1,3.2,2.8,3.3,3.3,2,3.1,1.4,2.7,18,13,14,14,1.6],
    #     ["Y","Wolf Knight Helm","head",4.2,4.4,3.8,4.4,4.4,3,4.2,2.2,3.8,23,16,17,17,2.9],
    #     ["Y","Wolf Knight Leggings","legs",5.1,6.6,5.5,6.6,6.6,4.2,6.4,2.8,5.6,25,15,19,19,4.6],
    #     ["Y","Wolnir's Crown","head",3.4,3.1,3.1,3.1,2.8,3.9,4.3,4.5,4.9,18,29,27,37,1.8],
    #     ["Y","Worker Garb","chest",4.2,4.4,6.8,5.5,5.5,6.6,5.2,5.2,10.9,34,68,43,60,2.3],
    #     ["Y","Worker Gloves","hands",1.4,0.9,1.5,1.1,1.1,1.2,0.9,0.9,2.3,11,22,14,19,0.8],
    #     ["Y","Worker Hat","head",2.3,1.9,2.7,2.2,2.2,3,2.5,2.5,4.5,16,31,20,27,1.1],
    #     ["Y","Worker Trousers","legs",2.9,2.9,4.2,3.5,3.5,4.5,3.7,3.7,7,23,44,28,39,1.6],
    #     ["N","Xanthous Crown","head","","","","","","","","","","","","","",""],
    #     ["Y","Xanthous Gloves","hands",2,1.8,2,1.2,1.8,2.7,1.2,2.9,2.6,15,22,22,18,0.8],
    #     ["Y","Xanthous Overcoat","chest",8.6,10.1,6.9,10.7,9.4,11.7,5.5,10,11.7,39,46,33,22,7.7],
    #     ["Y","Xanthous Trousers","legs",3.9,4.6,5.1,3.4,4.6,7,3.4,7.4,6.7,22,35,35,28,1.5]
    # ]
    with open("DS3_Armor_Calc.csv", "w", newline="") as file:
        csv.writer(file, delimiter=",").writerows(csvBackup)

with open("DS3_Armor_Calc.csv") as armor_csv:
    reader = csv.reader(armor_csv)
    data = [row for row in reader]

inputLayout = [
    [
        sg.Text("Weight Without Armor", size=(16)),
        sg.InputText(key="initWeight", size=(6)),
        sg.Text("Priority:"),
    ],
    [
        sg.Text("Weight Capacity", size=(16)),
        sg.InputText(key="capacity", size=(6)),
        sg.Checkbox("Standard", key="standardFlag", default=False, size=(6)),
        sg.Text("Multiplier:"),
        sg.InputText(key="standardCoef", size=(4), default_text="1"),
    ],
    [
        sg.Text("Target Percentage", size=(16)),
        sg.InputText(key="percent", default_text=69.9, size=(6)),
        sg.Checkbox("Strike", key="strikeFlag", default=False, size=(6)),
        sg.Text("Multiplier:"),
        sg.InputText(key="strikeCoef", size=(4), default_text="1"),
    ],
    [
        sg.Text("Armor:", size=23),
        sg.Checkbox("Slash", key="slashFlag", default=False, size=(6)),
        sg.Text("Multiplier:"),
        sg.InputText(key="slashCoef", size=(4), default_text="1"),
    ],
    [
        sg.Checkbox("Head", key="headFlag", default=True, size=(20)),
        sg.Checkbox("Thrust", key="thrustFlag", default=False, size=(6)),
        sg.Text("Multiplier:"),
        sg.InputText(key="thrustCoef", size=(4), default_text="1"),
    ],
    [
        sg.Checkbox("Chest", key="chestFlag", default=True, size=(20)),
        sg.Checkbox("Magic", key="magicFlag", default=False, size=(6)),
        sg.Text("Multiplier:"),
        sg.InputText(key="magicCoef", size=(4), default_text="1"),
    ],
    [
        sg.Checkbox("Hands", key="handsFlag", default=True, size=(20)),
        sg.Checkbox("Fire", key="fireFlag", default=False, size=(6)),
        sg.Text("Multiplier:"),
        sg.InputText(key="fireCoef", size=(4), default_text="1"),
    ],
    [
        sg.Checkbox("Legs", key="legsFlag", default=True, size=(20)),
        sg.Checkbox("Lightning", key="lightningFlag", default=False, size=(6)),
        sg.Text("Multiplier:"),
        sg.InputText(key="lightningCoef", size=(4), default_text="1"),
    ],
    [
        sg.Text("", size=23),
        sg.Checkbox("Dark", key="darkFlag", default=False, size=(6)),
        sg.Text("Multiplier:"),
        sg.InputText(key="darkCoef", size=(4), default_text="1"),
    ],
    [
        sg.Text("", size=23),
        # sg.Text("",size=1),
        # sg.Button("Inventory",size=17),
        # sg.Text("",size=1),
        sg.Checkbox("Bleed", key="bleedFlag", default=False, size=(6)),
        sg.Text("Multiplier:"),
        sg.InputText(key="bleedCoef", size=(4), default_text="0.1"),
    ],
    [
        sg.Text("", size=23),
        sg.Checkbox("Poison", key="poisonFlag", default=False, size=(6)),
        sg.Text("Multiplier:"),
        sg.InputText(key="poisonCoef", size=(4), default_text="0.1"),
    ],
    [
        sg.Text("", size=1),
        sg.Button("Calculate", size=17),
        sg.Text("", size=1),
        sg.Checkbox("Frost", key="frostFlag", default=False, size=(6)),
        sg.Text("Multiplier:"),
        sg.InputText(key="frostCoef", size=(4), default_text="0.1"),
    ],
    [
        sg.Text("", size=23),
        sg.Checkbox("Curse", key="curseFlag", default=False, size=(6)),
        sg.Text("Multiplier:"),
        sg.InputText(key="curseCoef", size=(4), default_text="0.1"),
    ],
    [
        sg.Text("", size=1),
        sg.Button("Exit", size=17),
        sg.Text("", size=1),
        sg.Checkbox("Poise", key="poiseFlag", default=False, size=(6)),
        sg.Text("Multiplier:"),
        sg.InputText(key="poiseCoef", size=(4), default_text="1"),
    ],
]


class LaunchValues:
    def __init__(
        args,
        weight,
        headFlag,
        chestFlag,
        handsFlag,
        legsFlag,
        standardFlag,
        standardCoef,
        strikeFlag,
        strikeCoef,
        slashFlag,
        slashCoef,
        thrustFlag,
        thrustCoef,
        magicFlag,
        magicCoef,
        fireFlag,
        fireCoef,
        lightningFlag,
        lightningCoef,
        darkFlag,
        darkCoef,
        bleedFlag,
        bleedCoef,
        poisonFlag,
        poisonCoef,
        frostFlag,
        frostCoef,
        curseFlag,
        curseCoef,
        poiseFlag,
        poiseCoef,
    ):
        args.weight = weight
        args.headFlag = headFlag
        args.chestFlag = chestFlag
        args.handsFlag = handsFlag
        args.legsFlag = legsFlag
        args.standardFlag = standardFlag
        args.standardCoef = standardCoef
        args.strikeFlag = strikeFlag
        args.strikeCoef = strikeCoef
        args.slashFlag = slashFlag
        args.slashCoef = slashCoef
        args.thrustFlag = thrustFlag
        args.thrustCoef = thrustCoef
        args.magicFlag = magicFlag
        args.magicCoef = magicCoef
        args.fireFlag = fireFlag
        args.fireCoef = fireCoef
        args.lightningFlag = lightningFlag
        args.lightningCoef = lightningCoef
        args.darkFlag = darkFlag
        args.darkCoef = darkCoef
        args.bleedFlag = bleedFlag
        args.bleedCoef = bleedCoef
        args.poisonFlag = poisonFlag
        args.poisonCoef = poisonCoef
        args.frostFlag = frostFlag
        args.frostCoef = frostCoef
        args.curseFlag = curseFlag
        args.curseCoef = curseCoef
        args.poiseFlag = poiseFlag
        args.poiseCoef = poiseCoef


class Armor:
    def __init__(
        armor,
        name,
        type,
        weight,
        standard,
        strike,
        slash,
        thrust,
        magic,
        fire,
        lightning,
        dark,
        bleed,
        poison,
        frost,
        curse,
        poise,
    ):
        armor.name = name
        armor.type = type
        armor.weight = weight
        armor.standard = standard
        armor.strike = strike
        armor.slash = slash
        armor.thrust = thrust
        armor.magic = magic
        armor.fire = fire
        armor.lightning = lightning
        armor.dark = dark
        armor.bleed = bleed
        armor.poison = poison
        armor.frost = frost
        armor.curse = curse
        armor.poise = poise


armorArray = []
for i in range(len(data)):
    if i > 0:  # ignore header
        if data[i][0] == "Y":  # only use active
            tempArmor = Armor(
                name=data[i][1],
                type=data[i][2],
                weight=float(data[i][3]),
                standard=float(data[i][4]),
                strike=float(data[i][5]),
                slash=float(data[i][6]),
                thrust=float(data[i][7]),
                magic=float(data[i][8]),
                fire=float(data[i][9]),
                lightning=float(data[i][10]),
                dark=float(data[i][11]),
                bleed=float(data[i][12]),
                poison=float(data[i][13]),
                frost=float(data[i][14]),
                curse=float(data[i][15]),
                poise=float(data[i][16]),
            )
            armorArray.append(tempArmor)


def calcArmorValue(launchValues, armorPiece):
    value = 0
    if launchValues.standardFlag:
        value = value + (armorPiece.standard * launchValues.standardCoef)
    if launchValues.strikeFlag:
        value = value + (armorPiece.strike * launchValues.strikeCoef)
    if launchValues.slashFlag:
        value = value + (armorPiece.slash * launchValues.slashCoef)
    if launchValues.thrustFlag:
        value = value + (armorPiece.thrust * launchValues.thrustCoef)
    if launchValues.magicFlag:
        value = value + (armorPiece.magic * launchValues.magicCoef)
    if launchValues.fireFlag:
        value = value + (armorPiece.fire * launchValues.fireCoef)
    if launchValues.lightningFlag:
        value = value + (armorPiece.lightning * launchValues.lightningCoef)
    if launchValues.darkFlag:
        value = value + (armorPiece.dark * launchValues.darkCoef)
    if launchValues.bleedFlag:
        value = value + (armorPiece.bleed * launchValues.bleedCoef)
    if launchValues.poisonFlag:
        value = value + (armorPiece.poison * launchValues.poisonCoef)
    if launchValues.frostFlag:
        value = value + (armorPiece.frost * launchValues.frostCoef)
    if launchValues.curseFlag:
        value = value + (armorPiece.curse * launchValues.curseCoef)
    if launchValues.poiseFlag:
        value = value + (armorPiece.poise * launchValues.poiseCoef)
    return value


def nextArmor(launchValues, remainingWeight, armorArray, armorPiece):
    armorValue = calcArmorValue(launchValues, armorPiece)
    popList = []
    nextPiece = False
    nextIncr = False
    armorRate = 0
    for i in range(len(armorArray)):
        nextValue = calcArmorValue(launchValues, armorArray[i])
        nextRate = nextValue / armorArray[i].weight
        if nextValue <= armorValue:
            popList.append(i)
        elif (armorArray[i].weight - armorPiece.weight) > remainingWeight:
            popList.append(i)
        elif nextRate > armorRate:
            nextPiece = armorArray[i]
            armorRate = nextRate
    for j in reversed(range(len(popList))):
        armorArray.pop(popList[j])
    if nextPiece:
        nextIncr = (nextValue - armorValue) / (nextPiece.weight - armorPiece.weight)
    return [nextPiece, nextIncr]


def initArmorArrays(launchValues, armorArray):
    helmArray = []
    chestArray = []
    gauntletArray = []
    legArray = []
    for i in range(len(armorArray)):
        match armorArray[i].type:
            case "head":
                if launchValues.headFlag:
                    helmArray.append(armorArray[i])
            case "chest":
                if launchValues.chestFlag:
                    chestArray.append(armorArray[i])
            case "hands":
                if launchValues.handsFlag:
                    gauntletArray.append(armorArray[i])
            case "legs":
                if launchValues.legsFlag:
                    legArray.append(armorArray[i])
            case _:
                raise Exception("Invalid Armor Type: " + armorArray[i])
    if not launchValues.headFlag:
        helmArray.append(
            Armor(
                name="",
                type="head",
                weight=float(0.0001),
                standard=float(0),
                strike=float(0),
                slash=float(0),
                thrust=float(0),
                magic=float(0),
                fire=float(0),
                lightning=float(0),
                dark=float(0),
                bleed=float(0),
                poison=float(0),
                frost=float(0),
                curse=float(0),
                poise=float(0),
            )
        )
    if not launchValues.chestFlag:
        chestArray.append(
            Armor(
                name="",
                type="chest",
                weight=float(0.0001),
                standard=float(0),
                strike=float(0),
                slash=float(0),
                thrust=float(0),
                magic=float(0),
                fire=float(0),
                lightning=float(0),
                dark=float(0),
                bleed=float(0),
                poison=float(0),
                frost=float(0),
                curse=float(0),
                poise=float(0),
            )
        )
    if not launchValues.handsFlag:
        gauntletArray.append(
            Armor(
                name="",
                type="hands",
                weight=float(0.0001),
                standard=float(0),
                strike=float(0),
                slash=float(0),
                thrust=float(0),
                magic=float(0),
                fire=float(0),
                lightning=float(0),
                dark=float(0),
                bleed=float(0),
                poison=float(0),
                frost=float(0),
                curse=float(0),
                poise=float(0),
            )
        )
    if not launchValues.legsFlag:
        legArray.append(
            Armor(
                name="",
                type="legs",
                weight=float(0.0001),
                standard=float(0),
                strike=float(0),
                slash=float(0),
                thrust=float(0),
                magic=float(0),
                fire=float(0),
                lightning=float(0),
                dark=float(0),
                bleed=float(0),
                poison=float(0),
                frost=float(0),
                curse=float(0),
                poise=float(0),
            )
        )
    return helmArray, chestArray, gauntletArray, legArray


def initOptimalArmor(launchValues, armorArray):
    helmArray, chestArray, gauntletArray, legArray = initArmorArrays(
        launchValues, armorArray
    )
    helmPiece = min(helmArray, key=attrgetter("weight"))
    chestPiece = min(chestArray, key=attrgetter("weight"))
    gauntletPiece = min(gauntletArray, key=attrgetter("weight"))
    legPiece = min(legArray, key=attrgetter("weight"))
    remainingWeight = launchValues.weight - (
        helmPiece.weight + chestPiece.weight + gauntletPiece.weight + legPiece.weight
    )
    nextHelm = False
    nextChest = False
    nextGauntlet = False
    nextLeg = False
    helmIncr = False
    chestIncr = False
    gauntletIncr = False
    legIncr = False
    if len(helmArray) > 0:
        nextHelm, helmIncr = nextArmor(
            launchValues, remainingWeight, helmArray, helmPiece
        )
    else:
        helmIncr = False
    if len(chestArray) > 0:
        nextChest, chestIncr = nextArmor(
            launchValues, remainingWeight, chestArray, chestPiece
        )
    else:
        chestIncr = False
    if len(gauntletArray) > 0:
        nextGauntlet, gauntletIncr = nextArmor(
            launchValues, remainingWeight, gauntletArray, gauntletPiece
        )
    else:
        gauntletIncr = False
    if len(legArray) > 0:
        nextLeg, legIncr = nextArmor(launchValues, remainingWeight, legArray, legPiece)
    else:
        legIncr = False
    return (
        helmArray,
        chestArray,
        gauntletArray,
        legArray,
        helmPiece,
        chestPiece,
        gauntletPiece,
        legPiece,
        nextHelm,
        nextChest,
        nextGauntlet,
        nextLeg,
        helmIncr,
        chestIncr,
        gauntletIncr,
        legIncr,
    )


def optimalArmor(
    launchValues,
    helmArray,
    chestArray,
    gauntletArray,
    legArray,
    helmPiece,
    chestPiece,
    gauntletPiece,
    legPiece,
    nextHelm,
    nextChest,
    nextGauntlet,
    nextLeg,
    helmIncr,
    chestIncr,
    gauntletIncr,
    legIncr,
):
    while (
        len(helmArray) > 0
        or len(chestArray) > 0
        or len(gauntletArray) > 0
        or len(legArray) > 0
    ):
        maxIncr = max(helmIncr, chestIncr, gauntletIncr, legIncr)
        if nextHelm and helmIncr == maxIncr:
            helmPiece = nextHelm
        elif nextChest and chestIncr == maxIncr:
            chestPiece = nextChest
        elif nextGauntlet and gauntletIncr == maxIncr:
            gauntletPiece = nextGauntlet
        elif nextLeg:
            legPiece = nextLeg
        remainingWeight = launchValues.weight - (
            helmPiece.weight
            + chestPiece.weight
            + gauntletPiece.weight
            + legPiece.weight
        )
        helmIncr = False
        chestIncr = False
        gauntletIncr = False
        legIncr = False
        nextHelm, helmIncr = nextArmor(
            launchValues, remainingWeight, helmArray, helmPiece
        )
        nextChest, chestIncr = nextArmor(
            launchValues, remainingWeight, chestArray, chestPiece
        )
        nextGauntlet, gauntletIncr = nextArmor(
            launchValues, remainingWeight, gauntletArray, gauntletPiece
        )
        nextLeg, legIncr = nextArmor(launchValues, remainingWeight, legArray, legPiece)
        # summary = [helmPiece.name, chestPiece.name, gauntletPiece.name, legPiece.name]
    return helmPiece, chestPiece, gauntletPiece, legPiece


def calcArmorSet(launchValues, armorArray):
    (
        helmArray,
        chestArray,
        gauntletArray,
        legArray,
        helmPiece,
        chestPiece,
        gauntletPiece,
        legPiece,
        nextHelm,
        nextChest,
        nextGauntlet,
        nextLeg,
        helmIncr,
        chestIncr,
        gauntletIncr,
        legIncr,
    ) = initOptimalArmor(launchValues, armorArray)
    helmPiece, chestPiece, gauntletPiece, legPiece = optimalArmor(
        launchValues,
        helmArray,
        chestArray,
        gauntletArray,
        legArray,
        helmPiece,
        chestPiece,
        gauntletPiece,
        legPiece,
        nextHelm,
        nextChest,
        nextGauntlet,
        nextLeg,
        helmIncr,
        chestIncr,
        gauntletIncr,
        legIncr,
    )
    return [helmPiece.name, chestPiece.name, gauntletPiece.name, legPiece.name]


# Create the window
launchWindow = sg.Window("DS3 Armor Calculator", inputLayout, size=(420, 440))


def getTargeWeight(launchArray):
    if launchArray["initWeight"] and launchArray["capacity"] and launchArray["percent"]:
        return float(launchArray["capacity"]) * float(
            launchArray["percent"]
        ) / 100 - float(launchArray["initWeight"])
    else:
        return 999


# Create an event loop
while True:
    launchEvent, launchArray = launchWindow.read()
    # Proceed if user closes window
    if launchEvent == "Exit" or launchEvent == sg.WIN_CLOSED:
        launchWindow.close()
        break
    # or presses the 'Calculate' button
    elif launchEvent == "Calculate":
        weight = getTargeWeight(launchArray)
        launchValues = LaunchValues(
            weight,
            launchArray["headFlag"],
            launchArray["chestFlag"],
            launchArray["handsFlag"],
            launchArray["legsFlag"],
            launchArray["standardFlag"],
            float(launchArray["standardCoef"]),
            launchArray["strikeFlag"],
            float(launchArray["strikeCoef"]),
            launchArray["slashFlag"],
            float(launchArray["slashCoef"]),
            launchArray["thrustFlag"],
            float(launchArray["thrustCoef"]),
            launchArray["magicFlag"],
            float(launchArray["magicCoef"]),
            launchArray["fireFlag"],
            float(launchArray["fireCoef"]),
            launchArray["lightningFlag"],
            float(launchArray["lightningCoef"]),
            launchArray["darkFlag"],
            float(launchArray["darkCoef"]),
            launchArray["bleedFlag"],
            float(launchArray["bleedCoef"]),
            launchArray["poisonFlag"],
            float(launchArray["poisonCoef"]),
            launchArray["frostFlag"],
            float(launchArray["frostCoef"]),
            launchArray["curseFlag"],
            float(launchArray["curseCoef"]),
            launchArray["poiseFlag"],
            float(launchArray["poiseCoef"]),
        )
        resultHelm, resultChest, resultGauntlet, resultLeg = calcArmorSet(
            launchValues, armorArray
        )
        resultLayout = [
            [sg.Text("Head:", size=(6)), sg.Text(resultHelm)],
            [sg.Text("Chest:", size=(6)), sg.Text(resultChest)],
            [sg.Text("Hands:", size=(6)), sg.Text(resultGauntlet)],
            [sg.Text("Legs:", size=(6)), sg.Text(resultLeg)],
            [sg.Button("OK")],
        ]
        resultWindow = sg.Window("Optimal Set", resultLayout)
        while True:
            resultEvent, resultValues = resultWindow.read()
            # End program if user closes window
            # or presses the 'OK' button
            if resultEvent == "OK" or resultEvent == sg.WIN_CLOSED:
                resultWindow.close()
                break
