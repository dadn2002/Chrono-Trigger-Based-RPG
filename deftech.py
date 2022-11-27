

# Just because it's too big

def tech(name, HP=None, MP=None, ATK=None, DFN=None, LVL=None, PWR=None, SPD=None, HIT=None, EVA=None, STM=None, MAG=None, MDF=None, RNG=None, mod=None):
    # The function that return the damage dealt by given 'name' tech with the statistcs of HP,MP,ATK,LVL,PWR,SPD,HIT,EVA
    # ,STM,MAG and the influence of RNG, DFN, EVA, MDF and mod in their effects.
    # This function returns an array: [damage, type[4], effects[255][3] where type is PHY or FIR/WAT/LIG/SHA
    # and effects is the set of debuffs : effects[0] = [name, level, duration]
    print('start the fucking function')

    # Crono main skills
    if name.lower() == 'cyclone':
        return [ATK*2.5*RNG, 'PHY', None]
    elif name.lower() == 'wind slash':
        return [(LVL+MAG)*4*RNG, 'PHY', None]
    elif name.lower() == 'lightning':
        return [(LVL+MAG)*4.8*RNG, 'LIG', None]
    elif name.lower() == 'cleave':
        return [ATK*4*RNG, 'PHY', None]
    elif name.lower() == 'lightning 2':
        return [(LVL+MAG)*5.6*RNG, 'LIG', None]
    elif name.lower() == 'raise':
        return [MAG*10*RNG, '', 'REV']
    elif name.lower() == 'frenzy':
        return [(ATK*1.81*RNG)*4, 'PHY', None]
    elif name.lower() == 'luminaire':
        return [(LVL+MAG)*20*RNG, 'LIG', 'PAR']

    # Lucca main skills
    elif name.lower() == 'flamethrower':
        return 0
    elif name.lower() == 'hypnowave':
        return 0
    elif name.lower() == 'fire':
        return 0
    elif name.lower() == 'napalm':
        return 0
    elif name.lower() == 'protect':
        return 0
    elif name.lower() == 'fire 2':
        return 0
    elif name.lower() == 'megaton bomb':
        return 0
    elif name.lower() == 'flare':
        return 0

    # Enemies Techs
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