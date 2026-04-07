const PptxGenJS = require("pptxgenjs");
// Note: Vercel uses Node 18+ with native fetch — no polyfill needed

const MAX_SLIDES = 20; // guard against oversized requests

const THEMES={
  professional:{dark:"1E2761",light:"FFFFFF",text:"1A2320",accent:"4472C4"},
  teal:        {dark:"1A9E8F",light:"FFFFFF",text:"1A2320",accent:"D95B2A"},
  warm:        {dark:"D95B2A",light:"FAFAF8",text:"1A2320",accent:"B04520"},
  minimal:     {dark:"1A2320",light:"FFFFFF",text:"1A2320",accent:"555555"},
  berry:       {dark:"6D2E46",light:"FAF7F4",text:"1A2320",accent:"A26769"},
  forest:      {dark:"2C5F2D",light:"FAFAF8",text:"1A2320",accent:"97BC62"}
};

// Returns a 6-char hex string, or null for transparent/near-transparent colours.
function toHex(c){
  if(!c||c==="transparent")return null;
  c=String(c).trim();
  if(/^#?[0-9A-Fa-f]{6}$/.test(c.replace("#","")))return c.replace("#","").toUpperCase();
  const m=c.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([\d.]+))?\s*\)/i);
  if(m){
    if(m[4]!==undefined&&parseFloat(m[4])<0.06)return null;
    return[m[1],m[2],m[3]].map(v=>("0"+parseInt(v).toString(16)).slice(-2)).join("").toUpperCase();
  }
  return"FFFFFF";
}

function getTheme(style,brandOn,brandColors){
  if(brandOn&&brandColors)return{dark:brandColors.primary.replace("#",""),light:"FFFFFF",text:brandColors.text.replace("#",""),accent:brandColors.accent.replace("#","")};
  return THEMES[style]||THEMES.professional;
}

async function fetchB64(url){
  try{const r=await fetch(url);const buf=await r.arrayBuffer();const ct=r.headers.get("content-type")||"image/jpeg";return ct+";base64,"+Buffer.from(buf).toString("base64");}catch(e){return null;}
}

async function prepImg(src){
  if(!src)return null;
  // PptxGenJS expects "mediaType;base64,<data>" without the leading "data:"
  if(src.startsWith("data:"))return src.slice(5);
  return await fetchB64(src);
}

async function getImgDims(d){
  try{
    const b64=d.includes(";base64,")?d.split(";base64,")[1]:d;
    const buf=Buffer.from(b64,"base64");
    // PNG: magic bytes 0x89 50 4E 47 at bytes 0-3
    if(buf[0]===0x89&&buf[1]===0x50&&buf[2]===0x4E&&buf[3]===0x47)
      return{w:buf.readUInt32BE(16),h:buf.readUInt32BE(20)};
    // JPEG: scan for SOF0/SOF2 markers
    let i=2;
    while(i<buf.length-8){
      if(buf[i]===0xFF&&(buf[i+1]===0xC0||buf[i+1]===0xC2))
        return{w:buf.readUInt16BE(i+7),h:buf.readUInt16BE(i+5)};
      const segLen=buf.readUInt16BE(i+2);
      if(segLen<2)break;
      i+=segLen+2;
    }
  }catch(e){}return null;
}

async function addLogo(s,pres,data,pos,wb,isCover){
  if(!data)return;
  const isBot=pos==="bottom-left",maxW=isCover?1.6:0.72,maxH=isCover?0.7:0.3;
  let fw=maxW,fh=maxH;
  const dims=await getImgDims(data);
  if(dims&&dims.w>0){const r=dims.w/dims.h;if(r>maxW/maxH){fw=maxW;fh=maxW/r;}else{fh=maxH;fw=maxH*r;}}
  const x=0.22,y=isBot?(5.625-fh-0.15):0.13;
  if(wb){const bh=Math.max(fh+0.15,isCover?0.9:0.48);s.addShape(pres.shapes.RECTANGLE,{x:0,y:isBot?5.625-bh:0,w:10,h:bh,fill:{color:"FFFFFF"},line:{color:"FFFFFF"}});}
  // PptxGenJS expects "mediaType;base64,<data>" (no leading "data:")
  const imgData=data.startsWith("data:")?data.slice(5):data;
  try{s.addImage({data:imgData,x,y,w:fw,h:fh});}catch(e){console.error("addLogo error:",e);}
}

async function buildPptx(slides,theme,pres,logoImg,logoPos,logoWb){
  const CW=960,CH=540,PW=10,PH=5.625;
  const ix=px=>Math.round(px/CW*PW*1000)/1000;
  const iy=px=>Math.round(px/CH*PH*1000)/1000;
  const clamp=(v,mn,mx)=>Math.max(mn,Math.min(mx,v));
  for(let si=0;si<slides.length;si++){
    const sd=slides[si];const s=pres.addSlide();
    const bgHex=toHex(sd.background||"#FFFFFF");
    s.background={color:bgHex||"FFFFFF"};
    await addLogo(s,pres,logoImg,logoPos||"top-left",logoWb,si===0);
    for(const el of(sd.elements||[])){
      const x=ix(clamp(el.left||0,-10,960));const y=iy(clamp(el.top||0,-10,540));
      const w=ix(clamp(el.width||100,2,960));const h=iy(clamp(el.height||20,1,540));
      if(el.type==="text"){
        const pt=clamp(Math.round((el.fontSize||18)*0.75),6,60);
        const colHex=toHex(el.color||"#1A2320");
        const col=colHex||"1A2320";
        const lh=el.lineHeight||1.35;
        const raw=el.text||"";
        const baseOpts={x,y,w,h,fontFace:el.fontFamily||"Calibri",bold:!!el.bold,italic:!!el.italic,
          fontSize:pt,color:col,align:el.align||"left",valign:"top",wrap:true,margin:[2,4,2,4]};
        if(raw.includes("\n")){
          const lines=raw.split("\n");
          s.addText(lines.map((ln,j)=>({text:ln,options:{fontSize:pt,fontFace:el.fontFamily||"Calibri",
            bold:!!el.bold,italic:!!el.italic,color:col,paraSpaceAfter:Math.round((lh-1)*pt*0.5),
            breakLine:j<lines.length-1}})),{x,y,w,h,valign:"top",wrap:true,margin:[2,4,2,4]});
        }else{s.addText(raw,baseOpts);}
      }else if(el.type==="image"){
        const d=await prepImg(el.src);
        if(d)try{s.addImage({data:d,x,y,w,h,sizing:{type:"cover",w,h}});}catch(e){}
      }else if(el.type==="shape"){
        if(w<0.01||h<0.01)continue;
        const fill=toHex(el.fill||"#cccccc");
        if(!fill)continue; // skip transparent/near-transparent shapes
        s.addShape(pres.shapes.RECTANGLE,{x,y,w,h,fill:{color:fill},line:{color:fill}});
      }else if(el.type==="circle"){
        if(w<0.01||h<0.01)continue;
        const cf_hex=el.fill&&el.fill!=="transparent"?toHex(el.fill):null;
        const cs_hex=el.stroke?toHex(el.stroke):null;
        if(cf_hex===null&&cs_hex===null)continue; // fully invisible
        const cf=cf_hex?{color:cf_hex}:{type:"none"};
        const cs=cs_hex?{color:cs_hex,pt:Math.max(1,Math.round((el.strokeWidth||2)*0.75))}:{type:"none"};
        s.addShape(pres.shapes.ELLIPSE,{x,y,w,h,fill:cf,line:cs});
      }
    }
  }
}

module.exports=async function handler(req,res){
  res.setHeader("Access-Control-Allow-Origin","*");
  res.setHeader("Access-Control-Allow-Methods","POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers","Content-Type");
  if(req.method==="OPTIONS")return res.status(200).end();
  if(req.method!=="POST")return res.status(405).json({error:"Method not allowed"});

  // Guard against missing body (e.g. malformed request)
  if(!req.body||typeof req.body!=="object")
    return res.status(400).json({error:"Request body missing or not JSON"});

  try{
    const{action,input,slideCount,style,title,slides,logoData,logoPos,logoWhiteBg,brandOn,brandColors}=req.body;
    const apiKey=process.env.ANTHROPIC_API_KEY;
    const model=process.env.CLAUDE_MODEL||"claude-sonnet-4-20250514";

    if(action==="outline"){
      // Cap slide count to prevent runaway API usage
      const count=Math.max(2,Math.min(MAX_SLIDES,parseInt(slideCount)||8));
      const prompt=`You are an expert presentation consultant. Create a ${count}-slide presentation.
USER INPUT:\n${input}

STRICT RULES:
- "heading": specific data-driven headline e.g. "Solar Costs Fell 89% in One Decade" NOT "Cost Trends"
- "bullets": EXACTLY 4 bullets, each a complete sentence 15-25 words with real numbers/data
- "paragraph": ALWAYS write 40-55 words of expert prose insight. Never omit.
- "subheading": 12-18 words for cover/closing slides only
- "stat": most striking number with units for stat layout
- "quote": realistic expert quote 20-30 words for quote layout
- "author": "Name, Title, Org" for quote layout
- First slide: layout "cover-center", dark:true
- Last slide: layout "closing", dark:true
- Mix layouts — use stat/quote/timeline/comparison/two-col to vary the deck

Return ONLY a raw JSON array of exactly ${count} objects. No markdown, no fences.
Each object must have ALL these exact fields:
{"title":"3-6 words","type":"title|content|data|quote|cta","dark":true,"layout":"cover-center|cover-split|cover-circle|cover-dark|default|two-col|three-col|title-body|quote|big-statement|agenda|closing|stat|three-stats|timeline|four-icons|two-icons|comparison|process|pyramid|img-right|img-left|img-full|img-top|two-images|three-images|img-mosaic|img-hero","heading":"specific headline","subheading":"","bullets":["Sentence one with data.","Sentence two.","Sentence three.","Sentence four."],"paragraph":"40-55 word expert insight paragraph.","stat":"","quote":"","author":"","imageKeyword":"3-5 words","speakerNote":"one sentence"}`;
      const r=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",
        headers:{"Content-Type":"application/json","x-api-key":apiKey,"anthropic-version":"2023-06-01"},
        body:JSON.stringify({model,max_tokens:8000,messages:[{role:"user",content:prompt}]})});
      const d=await r.json();
      if(!r.ok)return res.status(r.status).json({error:d.error?.message||"API error"});
      const text=d.content.map(b=>b.text||"").join("");
      const clean=text.replace(/```json|```/g,"").trim();
      const s=clean.indexOf("["),e=clean.lastIndexOf("]");
      if(s===-1||e===-1)return res.status(500).json({error:"No JSON array found in model response"});
      let outline;
      try{outline=JSON.parse(clean.slice(s,e+1));}
      catch(parseErr){return res.status(500).json({error:"Model returned invalid JSON: "+parseErr.message});}
      return res.status(200).json({outline});
    }

    if(action==="pptx"){
      const theme=getTheme(style,brandOn,brandColors);
      const pres=new PptxGenJS();
      pres.layout="LAYOUT_16x9";pres.title=title||"Presentation";
      // Normalise logo: keep full data URI; addLogo strips "data:" prefix internally
      let logoImg=null;
      if(logoData){
        logoImg=logoData.startsWith("data:")||logoData.includes(";base64,")
          ?logoData
          :await fetchB64(logoData);
      }
      try{
        await buildPptx(slides,theme,pres,logoImg,logoPos,logoWhiteBg);
      }catch(buildErr){
        console.error("buildPptx error:",buildErr);
        return res.status(500).json({error:"Failed to build slides: "+buildErr.message});
      }
      const buf=await pres.write({outputType:"nodebuffer"});
      const safe=((title||"presentation").trim()||"presentation").replace(/[^a-z0-9]/gi,"_");
      res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.presentationml.presentation");
      res.setHeader("Content-Disposition",`attachment; filename="${safe}.pptx"`);
      res.setHeader("Content-Length",buf.length);
      return res.status(200).send(buf);
    }

    return res.status(400).json({error:"Invalid action"});
  }catch(err){console.error(err);return res.status(500).json({error:err.message});}
};
