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

// Add logo with contain sizing to prevent stretching
async function addLogo(s,pres,logoImgData,logoPos,logoWhiteBg,isCover){
  if(!logoImgData)return;
  const isBottom=logoPos==='bottom-left';
  // Use small box and contain - pptxgenjs contain preserves aspect ratio
  const maxW=isCover?1.6:0.7;
  const maxH=isCover?0.6:0.25;
  const x=0.22;
  const y=isBottom?(5.625-maxH-0.15):0.13;
  if(logoWhiteBg){
    const barH=isCover?0.9:0.48;
    const barY=isBottom?5.625-barH:0;
    s.addShape(pres.shapes.RECTANGLE,{x:0,y:barY,w:10,h:barH,fill:{color:'FFFFFF'},line:{color:'FFFFFF'}});
  }
  try{
    // contain keeps aspect ratio, never stretches
    s.addImage({data:logoImgData,x,y,w:maxW,h:maxH,sizing:{type:'contain',w:maxW,h:maxH}});
  }catch(e){}
}

module.exports=async function handler(req,res){
  res.setHeader("Access-Control-Allow-Origin","*");
  res.setHeader("Access-Control-Allow-Methods","POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers","Content-Type");
  if(req.method==="OPTIONS")return res.status(200).end();
  if(req.method!=="POST")return res.status(405).json({error:"Method not allowed"});

  try{
    const{
      action,input,slideCount,style,title,outline,format,
      slideImages,slotImages,slideIcons,slideTemplates,slideDarkBg,
      logoData,logoPos,logoWhiteBg,brandOn,brandColors
    }=req.body;
    const apiKey=process.env.ANTHROPIC_API_KEY;

    // ── OUTLINE ──
    if(action==="outline"){
      const prompt=`You are a presentation design expert. Based on the input below, create a structured ${slideCount}-slide presentation outline.

${input}

IMPORTANT RULES:
- If the input contains document content or notes, read it carefully and adapt the ACTUAL content into slides — do NOT just summarise in generic bullet points. Extract key facts, data, arguments and insights.
- Mix content types: some slides can have bullet points, others a short paragraph of prose.

Return ONLY a raw JSON array with exactly ${slideCount} objects. No markdown, no code fences, no explanation — just the raw JSON array starting with [ and ending with ].

Each object must have:
- "title": string — short slide title, 3-7 words
- "type": string — one of: "title", "agenda", "content", "data", "quote", "cta", "conclusion"
- "bullets": array of 2-4 strings — concise bullet points using actual content (6-12 words each)
- "paragraph": string — a single prose paragraph of max 50 words expanding on the slide content (leave empty string "" if not needed)
- "speakerNote": string — one sentence of speaker guidance`;

      const r=await fetch("https://api.anthropic.com/v1/messages",{
        method:"POST",
        headers:{"Content-Type":"application/json","x-api-key":apiKey,"anthropic-version":"2023-06-01"},
        body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:3000,messages:[{role:"user",content:prompt}]})
      });
      const d=await r.json();
      if(!r.ok)return res.status(r.status).json({error:d.error?.message||"API error"});
      const text=d.content.map(b=>b.text||"").join("");
      // Robust JSON extraction — find the array even if there's surrounding text
      const cleaned=text.replace(/```json|```/g,"").trim();
      const jsonStart=cleaned.indexOf('[');
      const jsonEnd=cleaned.lastIndexOf(']');
      if(jsonStart===-1||jsonEnd===-1)return res.status(500).json({error:"Invalid outline response"});
      const parsed=JSON.parse(cleaned.slice(jsonStart,jsonEnd+1));
      return res.status(200).json({outline:parsed});
    }

    // ── PPTX ──
    if(action==="pptx"){
      const theme=getTheme(style,brandOn,brandColors);
      const pres=new PptxGenJS();
      pres.layout="LAYOUT_16x9";
      pres.title=title||"Presentation";
      const lastIdx=outline.length-1;

      // Prepare logo once
      let logoImg=null;
      if(logoData){
        logoImg=logoData.startsWith('data:')?logoData.replace('data:',''):await fetchAsBase64(logoData);
      }

      for(let i=0;i<outline.length;i++){
        const slide=outline[i];
        const s=pres.addSlide();
        const isCover=i===0;
        const isLast=i===lastIdx;
        const isSpecialSlide=isCover||isLast;
        // Per-slide dark background toggle
        const forceDark=slideDarkBg?.[i];
        const forceLight=forceDark===false;
        const defaultDark=slide.type==="title"||slide.type==="conclusion"||slide.type==="cta"||isCover||isLast;
        const isDark=forceDark===true?true:forceLight?false:defaultDark;

        const imgData=await prepImg(slideImages?.[i]);
        const tpl=slideTemplates?.[i]||'default';
        const slots=slotImages?.[i]||{};
        const icons=slideIcons?.[i]||{};

        s.background={color:isDark?theme.titleBg:theme.slideBg};

        // ── COVER / LAST SLIDE ──
        if(isSpecialSlide){
          if(tpl==='half-right'){
            // Half image right - text constrained to left 50%
            if(imgData){try{s.addImage({data:imgData,x:5,y:0,w:5,h:5.625,sizing:{type:'cover',w:5,h:5.625}});}catch(e){}}
            s.addShape(pres.shapes.RECTANGLE,{x:4.98,y:0,w:0.04,h:5.625,fill:{color:theme.accent},line:{color:theme.accent}});
            s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:0.1,h:5.625,fill:{color:theme.accent},line:{color:theme.accent}});
            await addLogo(s,pres,logoImg,logoPos||'top-left',logoWhiteBg,true);
            s.addText(slide.title,{x:0.5,y:1.6,w:4.3,h:1.5,fontSize:32,fontFace:"Calibri",bold:true,color:theme.titleText,align:"left"});
            if(slide.bullets?.[0])s.addText(slide.bullets[0],{x:0.5,y:3.3,w:4.3,h:0.7,fontSize:14,fontFace:"Calibri",color:theme.titleText,align:"left"});
          } else if(tpl==='circle-right'){
            await addLogo(s,pres,logoImg,logoPos||'top-left',logoWhiteBg,true);
            // Decorative rings
            const rings=[{r:4.2,op:'10'},{r:3.4,op:'15'},{r:2.6,op:'22'},{r:1.8,op:'35'}];
            rings.forEach(({r,op})=>{
              s.addShape(pres.shapes.OVAL,{x:10-r,y:(5.625-r*2)/2,w:r*2,h:r*2,
                fill:{color:'FFFFFF',transparency:100},
                line:{color:theme.accent,width:.5,transparency:parseInt(100-parseInt(op))}
              });
            });
            // Main circle clipped image
            if(imgData){try{s.addImage({data:imgData,x:5.8,y:0.3,w:4.0,h:5.0,sizing:{type:'cover',w:4.0,h:5.0},rounding:true});}catch(e){}}
            s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:0.1,h:5.625,fill:{color:theme.accent},line:{color:theme.accent}});
            s.addText(slide.title,{x:0.5,y:1.5,w:5.0,h:1.6,fontSize:32,fontFace:"Calibri",bold:true,color:theme.titleText,align:"left"});
            if(slide.bullets?.[0])s.addText(slide.bullets[0],{x:0.5,y:3.3,w:5.0,h:0.7,fontSize:14,fontFace:"Calibri",color:theme.titleText,align:"left"});
          } else {
            // Full bleed default
            if(imgData){try{s.addImage({data:imgData,x:0,y:0,w:10,h:5.625,sizing:{type:'cover',w:10,h:5.625},transparency:isDark?60:0});}catch(e){}}
            s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:0.1,h:5.625,fill:{color:theme.accent},line:{color:theme.accent}});
            await addLogo(s,pres,logoImg,logoPos||'top-left',logoWhiteBg,true);
            s.addText(slide.title,{x:0.6,y:1.6,w:8.8,h:1.6,fontSize:38,fontFace:"Calibri",bold:true,color:theme.titleText,align:"left"});
            if(slide.bullets?.[0])s.addText(slide.bullets[0],{x:0.6,y:3.4,w:8.5,h:0.7,fontSize:16,fontFace:"Calibri",color:theme.titleText,align:"left"});
          }

        // ── DARK CONTENT SLIDE ──
        }else if(isDark){
          if(imgData){try{s.addImage({data:imgData,x:0,y:0,w:10,h:5.625,sizing:{type:'cover',w:10,h:5.625},transparency:70});}catch(e){}}
          s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:0.08,h:5.625,fill:{color:theme.accent},line:{color:theme.accent}});
          await addLogo(s,pres,logoImg,logoPos||'top-left',logoWhiteBg,false);
          s.addText(slide.title,{x:0.6,y:1.6,w:8.8,h:1.4,fontSize:34,fontFace:"Calibri",bold:true,color:theme.titleText,align:"left"});
          if(slide.bullets?.[0])s.addText(slide.bullets[0],{x:0.6,y:3.2,w:7.5,h:0.7,fontSize:15,fontFace:"Calibri",color:theme.titleText,align:"left"});

        // ── 3 IMAGES ──
        }else if(tpl==='3images'){
          s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:10,h:0.06,fill:{color:theme.accent},line:{color:theme.accent}});
          await addLogo(s,pres,logoImg,logoPos||'top-left',logoWhiteBg,false);
          // Title at top
          const logoOffset=logoImg?0.45:0;
          s.addText(slide.title,{x:0.4,y:0.15+logoOffset,w:9.2,h:0.55,fontSize:20,fontFace:"Calibri",bold:true,color:theme.slideText,align:"left",margin:0});
          // 3 rectangular images side by side
          const imgW=3.0,imgH=2.8,imgY=0.95+logoOffset,gap=0.18;
          for(let k=0;k<3;k++){
            const imgX=0.3+k*(imgW+gap);
            const kImg=await prepImg(slots[k]);
            if(kImg){
              try{s.addImage({data:kImg,x:imgX,y:imgY,w:imgW,h:imgH,sizing:{type:'cover',w:imgW,h:imgH}});}
              catch(e){s.addShape(pres.shapes.RECTANGLE,{x:imgX,y:imgY,w:imgW,h:imgH,fill:{color:theme.accent+'33'},line:{color:theme.accent}});}
            }else{
              s.addShape(pres.shapes.RECTANGLE,{x:imgX,y:imgY,w:imgW,h:imgH,fill:{color:theme.accent+'22'},line:{color:theme.accent}});
              s.addText('Image '+(k+1),{x:imgX,y:imgY+imgH/2-0.2,w:imgW,h:0.4,fontSize:11,color:theme.accent,align:'center'});
            }
            // Title under image (bold)
            const textY=imgY+imgH+0.1;
            const bulletTitle=slide.bullets?.[k*2]||'';
            const bulletDesc=slide.bullets?.[k*2+1]||slide.bullets?.[k]||'';
            if(bulletTitle)s.addText(bulletTitle,{x:imgX,y:textY,w:imgW,h:0.3,fontSize:10,fontFace:"Calibri",bold:true,color:theme.slideText,align:'left'});
            if(bulletDesc)s.addText(bulletDesc,{x:imgX,y:textY+0.3,w:imgW,h:0.35,fontSize:9,fontFace:"Calibri",color:'666666',align:'left'});
          }

        // ── 2 or 4 ICONS ──
        }else if(tpl==='2icons'||tpl==='4icons'){
          const count=tpl==='2icons'?2:4;
          s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:10,h:0.06,fill:{color:theme.accent},line:{color:theme.accent}});
          await addLogo(s,pres,logoImg,logoPos||'top-left',logoWhiteBg,false);
          const logoOffset=logoImg?0.35:0;
          s.addText(slide.title,{x:0.5,y:0.15+logoOffset,w:9,h:0.6,fontSize:22,fontFace:"Calibri",bold:true,color:theme.slideText,align:"left",margin:0});
          const cellW=count===2?3.8:1.9;
          const startX=count===2?0.8:0.35;
          const gapX=count===2?1.0:0.35;
          const circleR=count===2?0.75:0.58;
          for(let k=0;k<count;k++){
            const cx=startX+k*(cellW+gapX);
            const ccx=cx+(cellW-circleR*2)/2;
            const circleY=1.0+logoOffset;
            // Circle outline only, primary colour
            s.addShape(pres.shapes.OVAL,{x:ccx,y:circleY,w:circleR*2,h:circleR*2,
              fill:{color:'FFFFFF',transparency:100},
              line:{color:theme.accent,width:1.5}
            });
            const iconInfo=icons[k];
            const slotImg=await prepImg(slots[k]);
            if(iconInfo){
              try{
                const svgRes=await fetch(iconInfo.svg);
                const svgText=await svgRes.text();
                // Colour the icon with accent colour
                const colouredSvg=svgText.replace(/fill="[^"]*"/g,`fill="#${theme.accent}"`).replace(/currentColor/g,`#${theme.accent}`);
                const svgB64='image/svg+xml;base64,'+Buffer.from(colouredSvg).toString('base64');
                const iSz=circleR*0.9;
                s.addImage({data:svgB64,x:ccx+(circleR*2-iSz)/2,y:circleY+(circleR*2-iSz)/2,w:iSz,h:iSz});
              }catch(e){}
            }else if(slotImg){
              try{const iSz=circleR*0.9;s.addImage({data:slotImg,x:ccx+(circleR*2-iSz)/2,y:circleY+(circleR*2-iSz)/2,w:iSz,h:iSz,sizing:{type:'contain',w:iSz,h:iSz}});}catch(e){}
            }
            const byY=circleY+circleR*2+0.18;
            s.addText((slide.bullets?.[k]||'Point '+(k+1)),{x:cx,y:byY,w:cellW,h:0.4,fontSize:count===2?14:11,fontFace:"Calibri",bold:true,color:theme.slideText,align:'center'});
            if(slide.bullets?.[k+count])s.addText(slide.bullets[k+count],{x:cx,y:byY+0.4,w:cellW,h:0.45,fontSize:9,color:'666666',align:'center'});
          }

        // ── FULL BLEED ──
        }else if(tpl==='fullbleed'){
          if(imgData){try{s.addImage({data:imgData,x:0,y:0,w:10,h:5.625,sizing:{type:'cover',w:10,h:5.625}});}catch(e){}}
          s.addShape(pres.shapes.RECTANGLE,{x:0,y:3.3,w:10,h:2.325,fill:{color:'000000'},line:{color:'000000'}});
          s.addShape(pres.shapes.RECTANGLE,{x:0,y:3.3,w:10,h:2.325,fill:{color:'000000',transparency:40},line:{color:'000000',transparency:40}});
          await addLogo(s,pres,logoImg,logoPos||'top-left',logoWhiteBg,false);
          s.addText(slide.title,{x:0.5,y:3.4,w:9,h:1.0,fontSize:26,fontFace:"Calibri",bold:true,color:'FFFFFF',align:'left'});
          if(slide.bullets?.[0])s.addText(slide.bullets[0],{x:0.5,y:4.5,w:9,h:0.6,fontSize:13,color:'FFFFFFCC',align:'left'});

        // ── TWO COLUMNS ──
        }else if(tpl==='two-col'){
          s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:10,h:0.06,fill:{color:theme.accent},line:{color:theme.accent}});
          await addLogo(s,pres,logoImg,logoPos||'top-left',logoWhiteBg,false);
          const logoOffset=logoImg?0.35:0;
          s.addText(slide.title,{x:0.5,y:0.15+logoOffset,w:9,h:0.65,fontSize:22,fontFace:"Calibri",bold:true,color:theme.slideText,align:"left",margin:0});
          s.addShape(pres.shapes.RECTANGLE,{x:0.5,y:0.9+logoOffset,w:0.9,h:0.04,fill:{color:theme.accent},line:{color:theme.accent}});
          s.addShape(pres.shapes.RECTANGLE,{x:5,y:1.0+logoOffset,w:0.04,h:4.3,fill:{color:theme.accent+'44'},line:{color:theme.accent+'44'}});
          const half=Math.ceil((slide.bullets||[]).length/2);
          const col1=slide.bullets?.slice(0,half)||[];
          const col2=slide.bullets?.slice(half)||[];
          const textY=1.05+logoOffset;
          if(col1.length)s.addText(col1.map((b,j)=>({text:b,options:{bullet:true,fontSize:13,fontFace:"Calibri",color:theme.slideText,paraSpaceAfter:6,breakLine:j<col1.length-1}})),{x:0.5,y:textY,w:4.3,h:4.3,valign:'top'});
          if(col2.length)s.addText(col2.map((b,j)=>({text:b,options:{bullet:true,fontSize:13,fontFace:"Calibri",color:theme.slideText,paraSpaceAfter:6,breakLine:j<col2.length-1}})),{x:5.2,y:textY,w:4.3,h:4.3,valign:'top'});

        // ── BIG STAT ──
        }else if(tpl==='stat'){
          s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:10,h:0.06,fill:{color:theme.accent},line:{color:theme.accent}});
          await addLogo(s,pres,logoImg,logoPos||'top-left',logoWhiteBg,false);
          const logoOffset=logoImg?0.35:0;
          s.addText(slide.title,{x:0.5,y:0.15+logoOffset,w:9,h:0.65,fontSize:22,fontFace:"Calibri",bold:true,color:theme.slideText,align:"center",margin:0});
          const stat=(slide.bullets||[])[0]||'100%';
          // Smaller stat text — was 72, now 48
          s.addText(stat,{x:1,y:1.1+logoOffset,w:8,h:1.6,fontSize:48,fontFace:"Calibri",bold:true,color:theme.accent,align:'center'});
          const rest=(slide.bullets||[]).slice(1);
          if(rest.length)s.addText(rest.join('  ·  '),{x:1,y:2.8+logoOffset,w:8,h:0.6,fontSize:13,color:'666666',align:'center'});
          // Image strip at bottom max ~250px (≈1.3in at 192dpi)
          const statImg=await prepImg(slideImages?.[i]);
          if(statImg){
            try{s.addImage({data:statImg,x:0,y:4.3,w:10,h:1.325,sizing:{type:'cover',w:10,h:1.325}});}catch(e){}
            // Semi-transparent overlay so it doesn't overpower
            s.addShape(pres.shapes.RECTANGLE,{x:0,y:4.3,w:10,h:1.325,fill:{color:theme.slideBg,transparency:30},line:{color:theme.slideBg,transparency:30}});
          }

        // ── QUOTE ──
        }else if(tpl==='quote'){
          s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:10,h:0.06,fill:{color:theme.accent},line:{color:theme.accent}});
          await addLogo(s,pres,logoImg,logoPos||'top-left',logoWhiteBg,false);
          const logoOffset=logoImg?0.35:0;
          const quote=(slide.bullets||[])[0]||slide.title;
          const author=(slide.bullets||[])[1]||'';
          // Large quote mark
          s.addText('\u201C',{x:0.5,y:0.4+logoOffset,w:1.5,h:1.2,fontSize:80,fontFace:'Georgia',color:theme.accent,align:'left'});
          // Quote text centred
          s.addText(quote,{x:0.8,y:1.3+logoOffset,w:8.4,h:2.5,fontSize:20,fontFace:'Calibri',italic:true,color:theme.slideText,align:'center',valign:'middle'});
          // Divider and author
          s.addShape(pres.shapes.RECTANGLE,{x:3.5,y:4.1+logoOffset,w:1.0,h:0.04,fill:{color:theme.accent},line:{color:theme.accent}});
          if(author)s.addText(author,{x:2,y:4.2+logoOffset,w:6,h:0.5,fontSize:13,fontFace:'Calibri',bold:true,color:theme.accent,align:'center'});

        // ── TIMELINE ──
        }else if(tpl==='timeline'){
          s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:10,h:0.06,fill:{color:theme.accent},line:{color:theme.accent}});
          await addLogo(s,pres,logoImg,logoPos||'top-left',logoWhiteBg,false);
          const logoOffset=logoImg?0.35:0;
          s.addText(slide.title,{x:0.5,y:0.15+logoOffset,w:9,h:0.65,fontSize:22,fontFace:'Calibri',bold:true,color:theme.slideText,align:'left',margin:0});
          const steps=(slide.bullets||[]).slice(0,4);
          const count=steps.length||4;
          const stepW=(9.0)/count;
          const lineY=2.2+logoOffset;
          // Horizontal connector line
          s.addShape(pres.shapes.RECTANGLE,{x:0.8,y:lineY+0.15,w:8.4,h:0.04,fill:{color:theme.accent,transparency:50},line:{color:theme.accent,transparency:50}});
          for(let k=0;k<count;k++){
            const cx=0.5+k*stepW+stepW/2;
            // Circle node
            s.addShape(pres.shapes.OVAL,{x:cx-0.22,y:lineY,w:0.44,h:0.44,fill:{color:theme.accent},line:{color:theme.accent}});
            s.addText(String(k+1),{x:cx-0.22,y:lineY,w:0.44,h:0.44,fontSize:11,fontFace:'Calibri',bold:true,color:'FFFFFF',align:'center',valign:'middle'});
            // Step text below
            if(steps[k])s.addText(steps[k],{x:cx-stepW/2+0.05,y:lineY+0.55,w:stepW-0.1,h:1.8,fontSize:11,fontFace:'Calibri',color:theme.slideText,align:'center',valign:'top'});
          }

        // ── DEFAULT (bullets + optional paragraph + optional image) ──
        }else{
          s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:10,h:0.06,fill:{color:theme.accent},line:{color:theme.accent}});
          await addLogo(s,pres,logoImg,logoPos||'top-left',logoWhiteBg,false);
          const logoOffset=logoImg?0.35:0;
          const textColor=isDark?'FFFFFF':theme.slideText;
          s.background={color:isDark?theme.titleBg:theme.slideBg};
          if(imgData){
            s.addText(slide.title,{x:0.5,y:0.15+logoOffset,w:5.6,h:0.7,fontSize:20,fontFace:"Calibri",bold:true,color:textColor,align:"left",margin:0});
            s.addShape(pres.shapes.RECTANGLE,{x:0.5,y:0.95+logoOffset,w:1.0,h:0.04,fill:{color:theme.accent},line:{color:theme.accent}});
            const bullH=slide.paragraph?2.4:3.8;
            if(slide.bullets?.length)s.addText(slide.bullets.map((b,j)=>({text:b,options:{bullet:true,fontSize:13,fontFace:"Calibri",color:textColor,paraSpaceAfter:6,breakLine:j<slide.bullets.length-1}})),{x:0.5,y:1.1+logoOffset,w:5.6,h:bullH,valign:'top'});
            if(slide.paragraph)s.addText(slide.paragraph,{x:0.5,y:1.1+logoOffset+bullH+0.1,w:5.6,h:0.85,fontSize:11,fontFace:"Calibri",color:textColor,italic:true});
            try{s.addImage({data:imgData,x:6.4,y:0.06,w:3.6,h:5.565,sizing:{type:'cover',w:3.6,h:5.565}});}catch(e){}
          }else{
            s.addText(slide.title,{x:0.5,y:0.15+logoOffset,w:9,h:0.7,fontSize:24,fontFace:"Calibri",bold:true,color:textColor,align:"left",margin:0});
            s.addShape(pres.shapes.RECTANGLE,{x:0.5,y:0.95+logoOffset,w:1.1,h:0.04,fill:{color:theme.accent},line:{color:theme.accent}});
            const bullH=slide.paragraph?2.6:4.0;
            if(slide.bullets?.length)s.addText(slide.bullets.map((b,j)=>({text:b,options:{bullet:true,fontSize:14,fontFace:"Calibri",color:textColor,paraSpaceAfter:7,breakLine:j<slide.bullets.length-1}})),{x:0.5,y:1.1+logoOffset,w:9,h:bullH,valign:'top'});
            if(slide.paragraph)s.addText(slide.paragraph,{x:0.5,y:1.1+logoOffset+bullH+0.15,w:9,h:0.85,fontSize:12,fontFace:"Calibri",color:textColor,italic:true});
          }
          if(slide.speakerNote)s.addNotes(slide.speakerNote);
        }
      }

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
