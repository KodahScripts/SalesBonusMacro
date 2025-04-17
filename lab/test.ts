432 - creates employees and attaches deals with baseline commission values 

90 - emp id [0]
     get Round([2] / 3)

3213 - emp id [0]
        value = [6] prior draw

NpsSheet - emp id [1]
           month surveys [5]
           month perc [8]
           90 day perc [23]

spiffs - emp id [0]
        value [7]

retropercentage = {
    if(unitcount >= 16) return 0.07
    if(unitcount >= 12 && unitcount < 16) return 0.04
    return 0
}

unitbonus = {
    if(unitcount >= 24) return 3000
    if(unitcount >= 20 unitcount < 24) return 2500
    if(unitcount >= 16 unitcount < 20) return 1500
    if(unitcount >= 12 unitcount < 16) return 750
    if(unitcount >= 10 unitcount < 12) return 375
    return 0
}

rollingMini = {
    if(unitcount >= 24) return 400
    if(unitcount >= 20 unitcount < 24) return 350
    if(unitcount >= 16 unitcount < 20) return 300
    if(unitcount >= 12 unitcount < 16) return 250
    return 0
}

npsforbonus = {
    if(monthlypercentage > 90daypercentage) return monthlypercentage
    return 90daypercentage
}

csioutcome = {
    `=IF(ISBLANK($A$2),"",IF(G2>$A$2+3%,"3P",IF(G2=$A$2,"A",IF(G2<$A$2,"B"))))`
}

results = {
    prior draw balance
    total commission amount
    total retro Commission
    total retro owed
    sum(total commission payout, total retro owed)
    total commission f&i
    above * -25%
    sum(above 2 lines)
    above * 5%
    if top saleman 500
    unit bonus
    `=IF(B5>=3,IF(B4="3P",L21*50,IF(B4="A",0,IF(B4="B",L21*-50))),0)`
    sum(unitbonus, CSI, top sales bonus)
    spiff total (from spiff sheet)
    SubmitEvent(Commission, total retro Commission, total f&i payout, total bonus, prior draw balance, spiff total)
}