var doc = app.activeDocument
var myMargin = doc.pages[0].marginPreferences.top; 
var pagesize = doc.documentPreferences; 
var height = pagesize.pageHeight; 
var width = pagesize.pageWidth;

var x0=doc.selection[0].geometricBounds[0];
var x1=doc.selection[0].geometricBounds[1];
var x2=doc.selection[0].geometricBounds[2];
var x3=doc.selection[0].geometricBounds[3];

doc.selection[0].fillColor = "None"
doc.selection[0].parentStory.fillColor = "Black"
doc.selection[0].geometricBounds =[x0,x1,x2+50,width/2-myMargin-2.5+x1]
doc.selection[0].parentStory.justification = Justification.CENTER_ALIGN
doc.selection[0].contents = doc.selection[0].contents.toUpperCase()
doc.selection[0].parentStory.capitalization=Capitalization.ALL_CAPS

nothings()

if(doc.selection[0].contents.search(/KORTING/g)!==-1){
    breakline(({findWhat:"(KORTING)(\\s)"}),({changeTo:"$1\\r"}))
}

else if(doc.selection[0].contents.search(/DE RÉDUCTION/g)!==-1){
    breakline(({findWhat:"(DE RÉDUCTION)(\\s)"}),({changeTo:"$1\\r"}))
}

else if(doc.selection[0].contents.search(/KAIKKI/g)!==-1){
    breakline(({findWhat:"(\\s)(\\-\\d+\\-\\d+\\%)"}),({changeTo:"\\r$2"}))
    nothings()
}

else if(doc.selection[0].contents.search(/KEDVEZMÉNY/g)!==-1){
    breakline(({findWhat:"(KEDVEZMÉNY)(\\s)"}),({changeTo:"$1\\r"}))
}

else{
    if(app.selection[0].contents.search(/\% \(LINE BREAK\)/g)!==-1||app.selection[0].contents.search(/\% \(LINEBREAK\)/g)!==-1){
        breakline(({findWhat:"(\\%)( \(LINE BREAK\))"}),({changeTo:"$1\\r"}))
    }
    else{
        breakline(({findWhat:"(\\%)(\\s)"}),({changeTo:"$1\\r"}))
    }
}

doc.selection[0].contents=doc.selection[0].contents.replace(/\(LINEBREAK\)|\(LINE BREAK\)/g,"\r")

nothings()

if(doc.selection[0].paragraphs.length==1){
    doc.selection[0].paragraphs[0].properties=({appliedFont:"Myriad Pro", fontStyle:"Black", pointSize:20, tracking:-15, leading:20, fillColor:"Black"})
}
else{
    if(doc.selection[0].contents.search(/KAIKKI/g)!==-1){
        doc.selection[0].paragraphs[0].properties=({appliedFont:"Myriad Pro", fontStyle:"Regular", pointSize:14, tracking:0, leading:14, fillColor:"Black"})
        doc.selection[0].paragraphs[1].properties=({appliedFont:"Myriad Pro", fontStyle:"Black", pointSize:28, tracking:-15, leading:28, fillColor:"JYSK_PC_400101XX"})
        if(doc.selection[0].paragraphs.length==3){
            doc.selection[0].paragraphs[2].properties=({appliedFont:"Myriad Pro", fontStyle:"Regular", pointSize:10, tracking:0, leading:10, fillColor:"Black"})
        }
    }
    else if(doc.selection[0].contents.search(/KEDVEZMÉNY|DE RÉDUCTION|PRIHRANITE|ÉCONOMISEZ/g)!==-1){
        doc.selection[0].paragraphs[0].properties=({appliedFont:"Myriad Pro", fontStyle:"Black", pointSize:25, tracking:-15, leading:25, fillColor:"JYSK_PC_400101XX"})
        doc.selection[0].paragraphs[1].properties=({appliedFont:"Myriad Pro", fontStyle:"Regular", pointSize:13, tracking:0, leading:13, fillColor:"Black"})
        if(doc.selection[0].paragraphs.length==3){
            doc.selection[0].paragraphs[2].properties=({appliedFont:"Myriad Pro", fontStyle:"Regular", pointSize:10, tracking:0, leading:10, fillColor:"Black"})
        }

    }
    else{
        doc.selection[0].paragraphs[0].properties=({appliedFont:"Myriad Pro", fontStyle:"Black", pointSize:28, tracking:-15, leading:28, fillColor:"JYSK_PC_400101XX"})
        doc.selection[0].paragraphs[1].properties=({appliedFont:"Myriad Pro", fontStyle:"Regular", pointSize:14, tracking:0, leading:14, fillColor:"Black"})
        if(doc.selection[0].paragraphs.length==3){
            doc.selection[0].paragraphs[2].properties=({appliedFont:"Myriad Pro", fontStyle:"Regular", pointSize:10, tracking:0, leading:10, fillColor:"Black"})
        }
    }
}
nothings()

app.findChangeGrepOptions.properties = ({includeFootnotes:true, kanaSensitive:true, widthSensitive:true});
app.findGrepPreferences.properties = ({findWhat:" BIS ZU | JUSQU’À | OP TIL | UPP TILL | OPPTIL | DO | AŽ | AKÁR | AŽ | DO | UP TO | ДО | DO | PÂNĂ LA | ΕΩΣ ΚΑΙ | TOT "});
app.changeGrepPreferences.properties = ({appliedFont:"Myriad Pro", fontStyle:"Regular"});
doc.selection[0].changeGrep();

nothings()

app.findChangeGrepOptions.properties = ({includeFootnotes:true, kanaSensitive:true, widthSensitive:true});
app.findGrepPreferences.properties = ({findWhat:"KAIKKI | ALT | TUTTO | TOTS | TOTA | TUTTA | TOTES | TUTTI | TUTTE | ALLE | TOUTES | TOUS | TODAS | TODA | TODOS | ВСИЧКИ | ALLA | WSZYSTKIE | VŠECHNY| VŠECHNA| VŠECHNA | VŠECHNO |MINDEN | VŠETKY| VŠETKY | VŠETOK | VŠETKO | VSE | VSO | VSI | ALL | BECb | ВСЮ | ВСЕ | ВСІ | SVIM | CJELOKUPNOM | CJELOKUPNOJ | CELOKUPNOM | TOT | TOATE | TOATĂ | ВСИЧКИ | ΟΛΟΥΣ | ΟΛΕΣ | ΟΛΑ | TOUS "});
app.changeGrepPreferences.properties = ({appliedFont:"Myriad Pro", fontStyle:"Black"});
doc.selection[0].changeGrep();

nothings()

if(doc.selection[0].contents.search(/KAIKKI/g)!==-1){
    app.findChangeGrepOptions.properties = ({includeFootnotes:true, kanaSensitive:true, widthSensitive:true});
    app.findGrepPreferences.properties = ({findWhat:"\\*"});
    app.changeGrepPreferences.properties = ({appliedFont:"Myriad Pro", fontStyle:"Regular"});
    app.changeGrepPreferences.position = Position.SUPERSCRIPT
    doc.selection[0].changeGrep();
}

doc.selection[0].fit(FitOptions.frameToContent)

function breakline(finded,changed){
    app.findChangeGrepOptions.properties = ({includeFootnotes:true, kanaSensitive:true, widthSensitive:true});
    app.findGrepPreferences.properties = finded;
    app.changeGrepPreferences.properties = changed;
    doc.selection[0].changeGrep();
}
function nothings(){
    app.findGrepPreferences = NothingEnum.NOTHING;
    app.changeGrepPreferences = NothingEnum.NOTHING;
}
