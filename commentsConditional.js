

if (value[elem]['light_purple']) {
    pObj.addLineBreak();
    pObj.addLineBreak();
    lightPurple = value[elem]['light_purple'];
    let separatorSymbol = '';
    //pObj.addText(lightPurple, { color: 'CC00CC', font_face: 'Calibri', font_size: 12 });
    if(lightPurple.substring(lightPurple.length - 1) === '/' || lightPurple.substring(lightPurple.length - 1) === '-'){
        separatorSymbol = lightPurple.substring(lightPurple.length - 1);
        lightPurple = lightPurple.slice(0,-1);
        pObj.addText(lightPurple, { color: 'CC00CC', font_face: 'Calibri', font_size: 12 });
        pObj.addText(separatorSymbol, {font_face: 'Calibri', font_size: 12 });
    }
}

if (value[elem]['blue']) {
    pObj.addLineBreak();
    pObj.addLineBreak();

    blue = value[elem]['blue'];
    pObj.addText(blue, { color: '366091', font_face: 'Calibri', font_size: 12 });
}

if (value[elem]['light_blue']) {
    pObj.addLineBreak()
    pObj.addLineBreak()
    lightBlue= value[elem]['light_blue'];
    pObj.addText(lightBlue, { color: '0070C0', font_face: 'Calibri', font_size: 12 });
}

if (value[elem]['red']) {
    red = value[elem]['red'];
    pObj.addText(' ' + red + ' ', { color: 'ff0000', font_face: 'Calibri', font_size: 12 });
}

if (value[elem]['blue_2']) {
    azulTwo = value[elem]['blue_2'];
    pObj.addText(azulTwo, { color: '366091', font_face: 'Calibri', font_size: 12 })
}

if (value[elem]['light_blue_2']) {
    lightBlueTwo = value[elem]['light_blue_2']
    pObj.addText(lightBlueTwo, { color: '0070C0', font_face: 'Calibri', font_size: 12 })
}

if (value[elem]['red_2']) {
    redTwo = value[elem]['red_2']
    pObj.addText(' ' + redTwo + ' ', { color: 'ff0000', font_face: 'Calibri', font_size: 12 })
}

if (value[elem]['dark_purple']) {
    pObj.addLineBreak();
    pObj.addLineBreak();
    darkPurple = value[elem]['dark_purple'];
    pObj.addText(' ' + darkPurple, { color: '800080', font_face: 'Calibri', font_size: 12 });
}

if (value[elem]['light_purple_2']) {
    pObj.addLineBreak()
    pObj.addLineBreak()
    lightPurpleTwo = value[elem]['light_purple_2']
    pObj.addLineBreak()
    pObj.addLineBreak()
    pObj.addText(' ' + lightPurpleTwo, { bold: true, color: 'CC00CC', font_face: 'Calibri', font_size: 12 })
}

if (value[elem]['dark_purple_2']) {
    pObj.addLineBreak();
    pObj.addLineBreak();
    darkPurpleTwo = value[elem]['dark_purple_2'];
    pObj.addText(' ' + darkPurpleTwo, { color: '800080', font_face: 'Calibri', font_size: 12 });
}

if (value[elem]['black_comment']) {
    pObj.addLineBreak()
    pObj.addLineBreak()
    blackComments = value[elem]['black_comment'];
    pObj.addText(blackComments, { font_face: 'Calibri', font_size: 12 });

    if (value[elem]['array bold']) {
        arrayBold.push(value[elem]['array bold']);
    }
}

if(value[elem]['black_group']) {
    pObj.addLineBreak()
    pObj.addLineBreak()
    blackGroup = value[elem]['black_group'];
    pObj.addText( blackGroup , {font_face: 'Calibri', font_size: 12 });
}

if(value[elem]['light_blue_group']) {
    pObj.addLineBreak()
    pObj.addLineBreak()
    lightBlueGroup = value[elem]['light_blue_group'];
    pObj.addText(lightBlueGroup, { color:'0070C0', font_face: 'Calibri', font_size: 12 });
}

if(value[elem]['red_group']) {
    pObj.addLineBreak()
    pObj.addLineBreak()
    redGroup = value[elem]['red_group'];
    pObj.addText(redGroup, { color: 'ff0000', font_face: 'Calibri', font_size: 12 });
}

if(value[elem]['light_purple_group']) {
    pObj.addLineBreak()
    pObj.addLineBreak()
    lightPurpleGroup = value[elem]['light_purple_group'];
    pObj.addText(lightPurpleGroup, { color: 'CC00CC', font_face: 'Calibri', font_size: 12 });
}

if(value[elem]['black_group2']) {
    pObj.addLineBreak()
    pObj.addLineBreak()
    blackGroup2 = value[elem]['black_group2'];
    pObj.addText( blackGroup2 , {font_face: 'Calibri', font_size: 12 });
}

if(value[elem]['light_blue_group2']) {
    pObj.addLineBreak()
    pObj.addLineBreak()
    lightBlueGroup2 = value[elem]['light_blue_group2'];
    pObj.addText(lightBlueGroup2, { color:'0070C0', font_face: 'Calibri', font_size: 12 });
}

if(value[elem]['red_group2']) {
    pObj.addLineBreak()
    pObj.addLineBreak()
    redGroup2 = value[elem]['red_group2'];
    pObj.addText(redGroup2, { color: 'ff0000', font_face: 'Calibri', font_size: 12 });
}

if(value[elem]['comments_group']){
    pObj.addLineBreak()
    pObj.addLineBreak()
    commentsGroup = value[elem]['comments_group'];
    pObj.addText(commentsGroup, { font_face: 'Calibri', font_size: 12 });
}

if (value[elem]['blue_3']) {
    pObj.addLineBreak()
    pObj.addLineBreak()
    blueThree = value[elem]['blue_3'];
    pObj.addText(blueThree, { color: '0070C0', font_face: 'Calibri', font_size: 12 });
}

if (value[elem]['red_3']) {
    pObj.addLineBreak();
    pObj.addLineBreak();
    redThree = value[elem]['red 3']
    pObj.addText(redThree, { color: 'ff0000', font_face: 'Calibri', font_size: 12 })
}

if (value[elem]['light_purple_three']) {
    pObj.addLineBreak();
    pObj.addLineBreak();
    lightPurpleThree = value[elem]['light_purple_three'];
    pObj.addText(lightPurpleThree, { color: 'CC00CC', font_face: 'Calibri', font_size: 12 });
}

if (value[elem]['dark_purple_three']) {
    pObj.addLineBreak();
    pObj.addLineBreak();
    darkPurpleThree = value[elem]['dark_purple_three'];
    pObj.addText(' ' + darkPurpleThree, { color: '800080', font_face: 'Calibri', font_size: 12 });
}
