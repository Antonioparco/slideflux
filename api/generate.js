const PptxGenJS = require("pptxgenjs");

module.exports = async function handler(req,res){
res.setHeader("Access-Control-Allow-Origin","*");
res.setHeader("Access-Control-Allow-Methods","POST,OPTIONS");
res.setHeader("Access-Control-Allow-Headers","Content-Type");

if(req.method==="OPTIONS") return res.status(200).end();
if(req.method!=="POST") return res.status(405).json({error:"Method not allowed"});

try{
const {action,slides,title} = req.body;

if(action==="pptx"){
const pres = new PptxGenJS();
pres.layout="LAYOUT_16X9";

slides.forEach(sl=>{
const s=pres.addSlide();

(sl.elements||[]).forEach(el=>{
if(el.type==="text"){
s.addText(el.text||"",{
x:el.left/96,
y:el.top/96,
w:el.width/96,
h:el.height/96,
fontSize:el.fontSize||16
});
}

if(el.type==="image" && el.src){
s.addImage({
path:el.src,
x:el.left/96,
y:el.top/96,
w:el.width/96,
h:el.height/96
});
}

if(el.type==="shape"){
s.addShape(pres.ShapeType.rect,{
x:el.left/96,
y:el.top/96,
w:el.width/96,
h:el.height/96,
fill:{color:"cccccc"}
});
}

if(el.type==="circle"){
s.addShape(pres.ShapeType.ellipse,{
x:el.left/96,
y:el.top/96,
w:el.width/96,
h:el.height/96,
fill:{color:"cccccc"}
});
}

if(el.type==="line"){
s.addShape(pres.ShapeType.line,{
x:el.left/96,
y:el.top/96,
w:Math.max(el.width/96,0.01),
h:Math.max(el.height/96,0.01),
line:{color:"000000",pt:2}
});
}
});
});

const buf=await pres.write({outputType:"nodebuffer"});
res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.presentationml.presentation");
res.setHeader("Content-Disposition",`attachment; filename="${title||"slides"}.pptx"`);
return res.send(buf);
}

return res.status(400).json({error:"Invalid action"});

}catch(e){
return res.status(500).json({error:e.message});
}
};