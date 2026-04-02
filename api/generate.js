const PptxGenJS = require("pptxgenjs");

const THEMES = {
  professional:{titleBg:"1E2761",titleText:"FFFFFF",slideBg:"FFFFFF",slideText:"1A2320",accent:"4472C4"},
  teal:{titleBg:"1A9E8F",titleText:"FFFFFF",slideBg:"FFFFFF",slideText:"1A2320",accent:"D95B2A"},
  warm:{titleBg:"D95B2A",titleText:"FFFFFF",slideBg:"FAFAF8",slideText:"1A2320",accent:"B04520"},
  minimal:{titleBg:"1A2320",titleText:"FFFFFF",slideBg:"FFFFFF",slideText:"1A2320",accent:"1A2320"},
  berry:{titleBg:"6D2E46",titleText:"FFFFFF",slideBg:"FAF7F4",slideText:"1A2320",accent:"A26769"},
  forest:{titleBg:"2C5F2D",titleText:"FFFFFF",slideBg:"FAFAF8",slideText:"1A2320",accent:"97BC62"}
};

async function fetchAsBase64(url){
  try{
    const res=await fetch(url);
    const buf=await res.arrayBuffer();
    const ct=res.headers.get('content-type')||'image/jpeg';
    return ct+';base64,'+Buffer.from(buf).toString('base64');
  }catch(e){return null;}
}

async function prepImg(d){
  if(!d)return null;
  if(d.startsWith('data:'))return d.replace('data:','');
  return await fetchAsBase64(d);
}

function getTheme(style,brandOn,brandColors){
  if(brandOn&&brandColors){
    return{
      titleBg:brandColors.primary.replace('#',''),
      titleText:'FFFFFF',
      slideBg:'FFFFFF',
      slideText:brandColors.text.replace('#',''),
      accent:brandColors.accent.replace('#','')
    };
  }
  return THEMES[style]||THEMES.professional;
}

async function getImgDimensions(dataStr){
  try{
    const b64=dataStr.includes(';base64,')?dataStr.split(';base64,')[1]:dataStr.split(',').pop();
    const buf=Buffer.from(b64,'base64');
    if(buf[1]===0x50&&buf[2]===0x4E&&buf[3]===0x47){
      return{w:buf.readUInt32BE(16),h:buf.readUInt32BE(20)};
    }
    let i=2;
    while(i<buf.length-8){
      if(buf[i]===0xFF&&(buf[i+1]===0xC0||buf[i+1]===0xC2)){
        return{w:buf.readUInt16BE(i+7),h:buf.readUInt16BE(i+5)};
      }
      i+=(i+2<buf.length?buf.readUInt16BE(i+2)+2:1);
    }
  }catch(e){}
  return null;
}

async function addLogo(s,pres,logoImgData,logoPos,logoWhiteBg,isCover){
  if(!logoImgData)return;
  const isBottom=logoPos==='bottom-left';
  const maxW=isCover?1.6:0.72;
  const maxH=isCover?0.7:0.3;
  let finalW=maxW,finalH=maxH;
  const dims=await getImgDimensions(logoImgData);
  if(dims&&dims.w>0&&dims.h>0){
    const ratio=dims.w/dims.h;
    if(ratio>maxW/maxH){finalW=maxW;finalH=maxW/ratio;}
    else{finalH=maxH;finalW=maxH*ratio;}
  }
  const x=0.22;
  const y=isBottom?(5.625-finalH-0.15):0.13;
  if(logoWhiteBg){
    const barH=Math.max(finalH+0.15,isCover?0.9:0.48);
    const barY=isBottom?5.625-barH:0;
    s.addShape(pres.shapes.RECTANGLE,{x:0,y:barY,w:10,h:barH,fill:{color:'FFFFFF'},line:{color:'FFFFFF'}});
  }
  try{s.addImage({data:logoImgData,x,y,w:finalW,h:finalH});}catch(e){}
}

// Build PPTX from slide objects — elements are in 960x540 coordinate space
async function buildPptxFromSlides(slides, theme, pres, logoImg, logoPos, logoWhiteBg){
  // Elements arrive already in 960x540 space (doExport normalizes them)
  // PPTX slide is 10" x 5.625" — so: inches = px/960*10
  const CW=960, CH=540, PW=10, PH=5.625;
  function toIn(px,axis){return Math.round((px/(axis==='x'?CW:CH)*(axis==='x'?PW:PH))*1000)/1000;}
  function clamp(v,mn,mx){return Math.max(mn,Math.min(mx,v));}
  function hexColor(c){
    if(!c)return 'FFFFFF';
    c=c.trim();
    if(c.startsWith('#'))return c.replace('#','').padEnd(6,'0').slice(0,6).toUpperCase();
    const m=c.match(/\d+/g);
    if(m&&m.length>=3)return [m[0],m[1],m[2]].map(v=>('0'+parseInt(v).toString(16)).slice(-2)).join('').toUpperCase();
    return 'FFFFFF';
  }

  for(let si=0;si<slides.length;si++){
    const slide=slides[si];
    const s=pres.addSlide();
    s.background={color:hexColor(slide.background||'#FFFFFF')};

    await addLogo(s,pres,logoImg,logoPos||'top-left',logoWhiteBg,si===0);

    const elements=slide.elements||[];
    for(const el of elements){
      // All coords are already in 960x540 space
      const x=toIn(clamp(el.left||0,-20,960),'x');
      const y=toIn(clamp(el.top||0,-20,540),'y');
      const w=toIn(clamp(el.width||100,4,960),'x');
      const h=toIn(clamp(el.height||20,2,540),'y');

      if(el.type==='text'){
        // Font size: elements are in 960-space pixels, convert to points (1pt = 1.33px at 96dpi)
        // But our canvas uses Arial at screen pixels, so 1px ≈ 0.75pt
        const rawPt=Math.round((el.fontSize||18)*0.75);
        const fontSize=clamp(rawPt,7,54);
        const color=hexColor(el.color||theme.slideText);
        const bold=!!el.bold;
        const italic=!!el.italic;
        const align=el.align||'left';
        const text=el.text||'';

        if(text.includes('\n')){
          // Multi-line: render as individual paragraphs
          const lines=text.split('\n').filter(l=>l.trim());
          s.addText(lines.map((line,j)=>({
            text:line,
            options:{fontSize,fontFace:'Calibri',bold,italic,color,
              paraSpaceAfter:2,breakLine:j<lines.length-1}
          })),{x,y,w,h,valign:'top',wrap:true});
        }else{
          s.addText(text,{x,y,w,h,fontSize,fontFace:'Calibri',
            bold,italic,color,align,valign:'middle',wrap:true,margin:2});
        }
      }else if(el.type==='image'&&el.src){
        const imgData=await prepImg(el.src);
        if(imgData){
          try{s.addImage({data:imgData,x,y,w,h,sizing:{type:'cover',w,h}});}catch(e){}
        }
      }else if(el.type==='shape'){
        if(w<0.01||h<0.01)continue; // skip invisible shapes
        const fill=hexColor(el.fill||'#cccccc');
        // Detect transparent/low-opacity fills and skip
        if(el.fill&&el.fill.includes('rgba')&&parseFloat((el.fill.match(/[\d.]+/g)||[])[3]||'1')<0.05)continue;
        s.addShape(pres.shapes.RECTANGLE,{x,y,w,h,fill:{color:fill},line:{color:fill}});
      }
    }
  }
}

module.exports=async function handler(req,res){
  res.setHeader("Access-Control-Allow-Origin","*");
  res.setHeader("Access-Control-Allow-Methods","POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers","Content-Type");
  if(req.method==="OPTIONS")return res.status(200).end();
  if(req.method!=="POST")return res.status(405).json({error:"Method not allowed"});

  try{
    const{action,input,slideCount,style,title,slides,
      logoData,logoPos,logoWhiteBg,brandOn,brandColors}=req.body;
    const apiKey=process.env.ANTHROPIC_API_KEY;

    // ── OUTLINE / SLIDE GENERATION ──
    if(action==="outline"){
      const prompt=`You are an expert presentation writer and consultant. Your job is to create a detailed, content-rich presentation that reads like it was written by a human expert — not a generic AI.

USER INPUT:
${input}

YOUR TASK:
Create exactly ${slideCount} slides. Each slide must be PACKED with real, specific, useful content. Think like a consultant who has researched this topic deeply.

STRICT CONTENT RULES:
- "heading": punchy, specific headline (not generic — e.g. "Solar Costs Fell 89% in a Decade" not "Cost Overview")
- "bullets": EXACTLY 4-5 bullets per slide. Each bullet must be a COMPLETE sentence or data point (15-25 words). Include real numbers, percentages, comparisons, examples wherever possible. No vague bullets like "Increased efficiency" — write "Solar panel efficiency improved from 15% to 23% between 2010-2024, reducing cost per watt by 60%."
- "paragraph": ALWAYS include this. Write 40-55 words of flowing prose that adds context beyond the bullets. This is the expert insight layer.
- "subheading": for title/conclusion slides write a compelling 12-18 word subtitle that sets the stage
- "stat": for stat slides pick the single most striking number with units (e.g. "£847bn" or "340%")
- "quote": for quote slides write a real or realistic expert quote (20-30 words) that supports the slide theme

LAYOUT RULES — pick the best layout for each slide's content:
- First slide: always "title-center" with dark:true
- Last slide: "title-center" with dark:true (call to action / conclusion)
- Use "stat" when a single number tells the whole story
- Use "quote" for credibility / social proof slides
- Use "two-col" when comparing two things or listing many points
- Use "timeline" for process, history, or steps
- Use "three-images" for showcasing products, case studies, or examples
- Use "two-icons" or "four-icons" for features, benefits, or pillars
- Use "fullbleed" for dramatic impact slides (market opportunity, vision)
- Use "default" (bullets) for most content slides

Return ONLY a raw JSON array of exactly ${slideCount} objects. No markdown, no explanation, no code fences.

Each object must have ALL these fields:
{
  "title": "short slide title 3-6 words",
  "type": "title|content|data|quote|cta|conclusion",
  "dark": true or false,
  "layout": "default|two-col|title-center|fullbleed|stat|quote|timeline|three-images|two-icons|four-icons",
  "heading": "specific compelling headline",
  "subheading": "supporting subtitle (use for title/conclusion slides)",
  "bullets": ["Complete sentence bullet with specific data point one","Complete sentence bullet with specific data point two","Complete sentence bullet with specific data point three","Complete sentence bullet with specific data point four"],
  "paragraph": "40-55 word expert insight paragraph that adds real context and depth beyond the bullets. Write this like a senior consultant summarising the key implication.",
  "stat": "the key number if layout is stat, empty string otherwise",
  "quote": "exact quote text if layout is quote, empty string otherwise",
  "author": "quote author name and title if layout is quote, empty string otherwise",
  "imageKeyword": "3-5 word specific image search term",
  "speakerNote": "one sentence of speaker guidance for this slide"
}`;

      const r=await fetch("https://api.anthropic.com/v1/messages",{
        method:"POST",
        headers:{"Content-Type":"application/json","x-api-key":apiKey,"anthropic-version":"2023-06-01"},
        body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:8000,messages:[{role:"user",content:prompt}]})
      });
      const d=await r.json();
      if(!r.ok)return res.status(r.status).json({error:d.error?.message||"API error"});
      const text=d.content.map(b=>b.text||"").join("");
      const cleaned=text.replace(/```json|```/g,"").trim();
      const jsonStart=cleaned.indexOf('[');
      const jsonEnd=cleaned.lastIndexOf(']');
      if(jsonStart===-1||jsonEnd===-1)return res.status(500).json({error:"Invalid response"});
      const outline=JSON.parse(cleaned.slice(jsonStart,jsonEnd+1));
      return res.status(200).json({outline});
    }

    // ── PPTX EXPORT ──
    if(action==="pptx"){
      const theme=getTheme(style,brandOn,brandColors);
      const pres=new PptxGenJS();
      pres.layout="LAYOUT_16x9";
      pres.title=title||"Presentation";

      let logoImg=null;
      if(logoData){
        logoImg=logoData.startsWith('data:')?logoData.replace('data:',''):await fetchAsBase64(logoData);
      }

      await buildPptxFromSlides(slides,theme,pres,logoImg,logoPos,logoWhiteBg);

      const buffer=await pres.write({outputType:"nodebuffer"});
      const safeName=(title||"presentation").replace(/[^a-z0-9]/gi,"_");
      res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.presentationml.presentation");
      res.setHeader("Content-Disposition",`attachment; filename="${safeName}.pptx"`);
      res.setHeader("Content-Length",buffer.length);
      return res.status(200).send(buffer);
    }

    return res.status(400).json({error:"Invalid action"});

  }catch(err){
    console.error("Error:",err);
    return res.status(500).json({error:err.message});
  }
};
