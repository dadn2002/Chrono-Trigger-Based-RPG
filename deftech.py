

# Just because it's too big
import openpyxl
from pathlib import Path


def grouptechs(name):
    # Return True if given tech is not target skill
    # print('testing the fucking function')
    if name.lower() == 'lightning 2':
        return True
    elif name.lower() == 'luminaire':
        return True
    elif name.lower() == 'fire 2':
        return True
    elif name.lower() == 'cyclone':
        return True
    elif name.lower() == 'megaton bomb':
        return True
    elif name.lower() == 'flare':
        return True
    elif name.lower() == 'smoke bomb':
        return True
    elif name.lower() == 'crushing blow':
        return True
    elif name.lower() == 'ink':
        return True
    elif name.lower() == 'sonic wave':
        return True
    elif name.lower() == 'earthquake':
        return True
    elif name.lower() == 'luminaire':
        return True
    elif name.lower() == 'luminaire':
        return True
    elif name.lower() == 'luminaire':
        return True
    elif name.lower() == 'luminaire':
        return True

    return False


def learnTech(name, player):
    TechListabc = Path("TechList.xlsx")
    TechListabcobj = openpyxl.load_workbook(TechListabc, data_only=True)
    TechList = TechListabcobj.active

    datacolumn = ''
    tpcolumn = ''
    if player.lower() == 'crono':
        datacolumn = 'C'
        tpcolumn = 'K'
    elif player.lower() == 'lucca':
        datacolumn = 'D'
        tpcolumn = 'L'
    elif player.lower() == 'marle':
        datacolumn = 'E'
        tpcolumn = 'M'
    elif player.lower() == 'frog':
        datacolumn = 'F'
        tpcolumn = 'N'
    elif player.lower() == 'robo':
        datacolumn = 'G'
        tpcolumn = 'O'
    elif player.lower() == 'ayla':
        datacolumn = 'H'
        tpcolumn = 'P'
    elif player.lower() == 'magus':
        datacolumn = 'I'
        tpcolumn = 'Q'
    else:
        print('TechList Error', player, 'is not a valid character')
        return 0
    i = 0
    testvariable = False
    testi = 0
    while TechList[datacolumn + str(4 + i)].value is not None:
        testi = i
        if str(TechList[datacolumn + str(4 + i)].value).lower() == name.lower():
            print('Already know tech')
            return 0
        i = i + 1
    if testvariable is False:
        tpcost = tech(name.lower(), mod='unlock')
        if tpcost is None:
            print(name, 'is not a valid tech')
            return 0
        if int(TechList[tpcolumn + '4'].value) - tpcost < 0:
            print('Not enough tp to learn', name)
            return 0
        TechList[tpcolumn + '4'] = int(TechList[tpcolumn + '4'].value) - tpcost
        TechList[datacolumn + str(4 + testi + 1)] = name
        print(player, 'learned', name)
    TechListabcobj.save(TechListabc)


def tech(name, HP=None, MP=None, ATK=None, DFN=None, LVL=None, PWR=None, SPD=None, HIT=None, EVA=None, STM=None, MAG=None, MDF=None, RNG=None, mod=''):
    # The function that return the damage dealt by given 'name' tech with the statistcs of HP,MP,ATK,LVL,PWR,SPD,HIT,EVA
    # ,STM,MAG and the influence of RNG, DFN, EVA, MDF and mod in their effects.
    # This function returns an array: [damage, type[4], effects[255][3] where type is PHY or FIR/WAT/LIG/SHA
    # and effects is the set of debuffs : effects[0] = [name, level, duration]
    # print('start the fucking function')

    # Crono main skills
    if name.lower() == 'cyclone':
        if mod == 'unlock':
            return 5
        elif mod == 'cost':
            return 2
        return [ATK*2.5*RNG, 'PHY', None]
    elif name.lower() == 'wind slash':
        if mod == 'unlock':
            return 50
        elif mod == 'cost':
            return 2
        return [(LVL+MAG)*4*RNG, 'PHY', None]
    elif name.lower() == 'lightning':
        if mod == 'unlock':
            return -1
        elif mod == 'cost':
            return 5
        return [(LVL+MAG)*4.8*RNG, 'LIG', None]
    elif name.lower() == 'cleave':
        if mod == 'unlock':
            return 150
        elif mod == 'cost':
            return 8
        return [ATK*4*RNG, 'PHY', None]
    elif name.lower() == 'lightning 2':
        if mod == 'unlock':
            return 400
        elif mod == 'cost':
            return 12
        return [(LVL+MAG)*5.6*RNG, 'LIG', None]
    elif name.lower() == 'raise':
        if mod == 'unlock':
            return 300
        elif mod == 'cost':
            return 10
        return [MAG*10*RNG, '', 'REV']
    elif name.lower() == 'frenzy':
        if mod == 'unlock':
            return 600
        elif mod == 'cost':
            return 12
        return [(ATK*1.81*RNG)*4, 'PHY', None]
    elif name.lower() == 'luminaire':
        if mod == 'unlock':
            return 800
        elif mod == 'cost':
            return 20
        return [(LVL+MAG)*20*RNG, 'LIG', 'PAR']

    # Lucca main skills
    elif name.lower() == 'flamethrower':
        if mod == 'unlock':
            return 5
        return 0
    elif name.lower() == 'hypnowave':
        if mod == 'unlock':
            return 40
        return 0
    elif name.lower() == 'fire':
        if mod == 'unlock':
            return -1
        return 0
    elif name.lower() == 'napalm':
        if mod == 'unlock':
            return 120
        return 0
    elif name.lower() == 'protect':
        if mod == 'unlock':
            return 200
        return 0
    elif name.lower() == 'fire 2':
        if mod == 'unlock':
            return 400
        return 0
    elif name.lower() == 'megaton bomb':
        if mod == 'unlock':
            return 500
        return 0
    elif name.lower() == 'flare':
        if mod == 'unlock':
            return 800
        return 0

    # Enemies Techs
    elif name.lower() == 'punching glove':
        return [ATK*1.5*RNG, 'PHY', None]
    elif name.lower() == 'charge':
        return [(ATK*1/4+HP*3/4)*1.5*RNG, 'PHY', None]
    elif name.lower() == 'fing flap':
        return [(ATK*1/4+HP*1/2)*SPD/3*RNG, 'PHY', None]
    elif name.lower() == 'ding-a-ling':
        return [0, '', 'SLP']
    elif name.lower() == 'strike':
        return [ATK*1.8*RNG, 'PHY', None]
    elif name.lower() == 'smash':
        return [ATK*1.5*RNG, 'PHY', 'STN']
    elif name.lower() == 'smoke bomb':
        return [MAG*1.2*RNG, 'SHA', 'BLD']
    elif name.lower() == 'hammer':
        return [(ATK+HP/10)*RNG, 'PHY', 'STN']
    elif name.lower() == 'crushing blow':
        return [(ATK/2+HP/20)*RNG, 'PHY', 'STN']
    elif name.lower() == 'stab':
        return [(ATK*2*SPD/2)*RNG, 'PHY', '']
    elif name.lower() == 'flame':
        return [MAG*1.5*RNG, 'FIR', '']
    elif name.lower() == 'clamp':
        return [HP/5*RNG, 'PHY', '']
    elif name.lower() == 'ink':
        return [0, 'SHA', 'BLD']
    elif name.lower() == 'osmose 1':
        return [5, '', 'MP']
    elif name.lower() == 'gore':
        return [ATK*1.5*RNG, 'PHY', 'STN']
    elif name.lower() == 'spin jump':
        return [HP/5*RNG, 'PHY', '']
    elif name.lower() == 'sonic wave':
        return [0, 'SHA', 'SLP']
    elif name.lower() == 'grudge':
        return [ATK*3*RNG, 'PHY', '']

    # Yakra Techs
    elif name.lower() == 'needlespin':
        return [(ATK-5)*RNG, 'PHY', '']
    elif name.lower() == 'earthquake':
        return [(ATK-20)*RNG, 'PHY', '']
    elif name.lower() == 'spitting rock':
        return [(ATK/5+HP/25)*RNG, 'PHY', '']
    elif name.lower() == 'pounce':
        return [HP/10*RNG, 'PHY', '']

    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    elif name.lower() == 'cyclone':
        return 0
    else:
        return None