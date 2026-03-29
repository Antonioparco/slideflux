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

async function addLogo(s,pres,logoImgData,logoPos,logoWhiteBg,isCover){
  if(!logoImgData)return;
  const isBottom=logoPos==='bottom-left';
  if(logoWhiteBg){
    const barH=isCover?1.1:0.55;
    const y=isBottom?5.625-barH:0;
    s.addShape(pres.shapes.RECTANGLE,{x:0,y,w:10,h:barH,fill:{color:'FFFFFF'},line:{color:'FFFFFF'}});
  }
  const w=isCover?1.6:0.8,h=isCover?0.65:0.32;
  const x=0.2;
  const y=isBottom?(5.625-h-0.1):0.1;
  try{s.addImage({data:logoImgData,x,y,w,h,sizing:{type:'contain',w,h}});}catch(e){}
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
      slideImages,slotImages,slideIcons,slideTemplates,
      logoData,logoPos,logoWhiteBg,brandOn,brandColors
    }=req.body;
    const apiKey=process.env.ANTHROPIC_API_KEY;

    // ── OUTLINE ──
    if(action==="outline"){
      const prompt=`You are a presentation design expert. Based on the input below, create a structured ${slideCount}-slide outline.\n\n${input}\n\nReturn ONLY a raw JSON array with exactly ${slideCount} objects. Each object:\n- "title": short slide title (3-7 words)\n- "type": one of "title","agenda","content","data","quote","cta","conclusion"\n- "bullets": 2-4 concise bullets (5-10 words each)\n- "speakerNote": one sentence of guidance\n\nNo markdown, no explanation, raw JSON array only.`;
      const r=await fetch("https://api.anthropic.com/v1/messages",{
        method:"POST",
        headers:{"Content-Type":"application/json","x-api-key":apiKey,"anthropic-version":"2023-06-01"},
        body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:1000,messages:[{role:"user",content:prompt}]})
      });
      const d=await r.json();
      if(!r.ok)return res.status(r.status).json({error:d.error?.message||"API error"});
      const text=d.content.map(b=>b.text||"").join("");
      return res.status(200).json({outline:JSON.parse(text.replace(/```json|```/g,"").trim())});
    }

    // ── PPTX ──
    if(action==="pptx"){
      const theme=getTheme(style,brandOn,brandColors);
      const pres=new PptxGenJS();
      pres.layout="LAYOUT_16x9";
      pres.title=title||"Presentation";

      // Prepare logo once
      let logoImg=null;
      if(logoData){
        logoImg=logoData.startsWith('data:')?logoData.replace('data:',''):await fetchAsBase64(logoData);
      }

      for(let i=0;i<outline.length;i++){
        const slide=outline[i];
        const s=pres.addSlide();
        const isDark=slide.type==="title"||slide.type==="conclusion"||slide.type==="cta"||i===0;
        const imgData=await prepImg(slideImages?.[i]);
        const tpl=slideTemplates?.[i]||'default';
        const isCover=i===0;
        const slots=slotImages?.[i]||{};
        const icons=slideIcons?.[i]||{};

        s.background={color:isDark?theme.titleBg:theme.slideBg};

        // ── DARK SLIDE ──
        if(isDark){
          if(imgData){try{s.addImage({data:imgData,x:0,y:0,w:10,h:5.625,transparency:70});}catch(e){}}
          s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:0.08,h:5.625,fill:{color:theme.accent},line:{color:theme.accent}});
          await addLogo(s,pres,logoImg,logoPos||'top-left',logoWhiteBg,isCover);
          s.addText(slide.title,{x:0.6,y:1.8,w:8.8,h:1.4,fontSize:36,fontFace:"Calibri",bold:true,color:theme.titleText,align:"left"});
          if(slide.bullets?.[0])s.addText(slide.bullets[0],{x:0.6,y:3.4,w:7.5,h:0.6,fontSize:15,fontFace:"Calibri",color:theme.titleText,align:"left"});

        // ── 3 IMAGES ──
        } else if(tpl==='3images'){
          s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:10,h:0.06,fill:{color:theme.accent},line:{color:theme.accent}});
          await addLogo(s,pres,logoImg,logoPos||'top-left',logoWhiteBg,false);
          s.addText(slide.title,{x:0.5,y:0.15,w:9,h:0.65,fontSize:22,fontFace:"Calibri",bold:true,color:theme.slideText,align:"left",margin:0});
          const imgW=2.9,imgH=3.2,imgY=1.0,gap=0.15;
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
            if(slide.bullets?.[k])s.addText(slide.bullets[k],{x:imgX,y:imgY+imgH+0.05,w:imgW,h:0.35,fontSize:9,color:theme.slideText,align:'center'});
          }

        // ── 2 or 4 ICONS ──
        } else if(tpl==='2icons'||tpl==='4icons'){
          const count=tpl==='2icons'?2:4;
          s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:10,h:0.06,fill:{color:theme.accent},line:{color:theme.accent}});
          await addLogo(s,pres,logoImg,logoPos||'top-left',logoWhiteBg,false);
          s.addText(slide.title,{x:0.5,y:0.15,w:9,h:0.65,fontSize:22,fontFace:"Calibri",bold:true,color:theme.slideText,align:"left",margin:0});
          const cellW=count===2?3.6:2.0;
          const startX=count===2?0.8:0.4;
          const gapX=count===2?1.2:0.4;
          const circleR=count===2?0.7:0.55;
          for(let k=0;k<count;k++){
            const cx=startX+k*(cellW+gapX);
            const ccx=cx+(cellW-circleR*2)/2;
            s.addShape(pres.shapes.OVAL,{x:ccx,y:1.1,w:circleR*2,h:circleR*2,fill:{color:theme.accent+'33'},line:{color:theme.accent}});
            const iconInfo=icons[k];
            const slotImg=await prepImg(slots[k]);
            if(iconInfo){
              try{
                const svgRes=await fetch(iconInfo.svg);
                const svgText=await svgRes.text();
                const svgB64='image/svg+xml;base64,'+Buffer.from(svgText).toString('base64');
                const iSz=circleR;
                s.addImage({data:svgB64,x:ccx+(circleR*2-iSz)/2,y:1.1+(circleR*2-iSz)/2,w:iSz,h:iSz});
              }catch(e){}
            }else if(slotImg){
              try{const iSz=circleR;s.addImage({data:slotImg,x:ccx+(circleR*2-iSz)/2,y:1.1+(circleR*2-iSz)/2,w:iSz,h:iSz,sizing:{type:'contain',w:iSz,h:iSz}});}catch(e){}
            }
            const byY=1.1+circleR*2+0.15;
            s.addText((slide.bullets?.[k]||'Point '+(k+1)),{x:cx,y:byY,w:cellW,h:0.45,fontSize:count===2?14:11,fontFace:"Calibri",bold:true,color:theme.slideText,align:'center'});
            if(slide.bullets?.[k+count])s.addText(slide.bullets[k+count],{x:cx,y:byY+0.45,w:cellW,h:0.45,fontSize:9,color:'666666',align:'center'});
          }

        // ── FULL BLEED ──
        } else if(tpl==='fullbleed'){
          if(imgData){try{s.addImage({data:imgData,x:0,y:0,w:10,h:5.625,sizing:{type:'cover',w:10,h:5.625}});}catch(e){}}
          s.addShape(pres.shapes.RECTANGLE,{x:0,y:3.3,w:10,h:2.325,fill:{color:'000000'},line:{color:'000000'}});
          s.addShape(pres.shapes.RECTANGLE,{x:0,y:3.3,w:10,h:2.325,fill:{color:'000000',transparency:45},line:{color:'000000',transparency:45}});
          await addLogo(s,pres,logoImg,logoPos||'top-left',logoWhiteBg,false);
          s.addText(slide.title,{x:0.5,y:3.4,w:9,h:1.0,fontSize:26,fontFace:"Calibri",bold:true,color:'FFFFFF',align:'left'});
          if(slide.bullets?.[0])s.addText(slide.bullets[0],{x:0.5,y:4.5,w:9,h:0.6,fontSize:13,color:'FFFFFFCC',align:'left'});

        // ── TWO COLUMNS ──
        } else if(tpl==='two-col'){
          s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:10,h:0.06,fill:{color:theme.accent},line:{color:theme.accent}});
          await addLogo(s,pres,logoImg,logoPos||'top-left',logoWhiteBg,false);
          s.addText(slide.title,{x:0.5,y:0.15,w:9,h:0.7,fontSize:22,fontFace:"Calibri",bold:true,color:theme.slideText,align:"left",margin:0});
          s.addShape(pres.shapes.RECTANGLE,{x:0.5,y:0.95,w:0.9,h:0.04,fill:{color:theme.accent},line:{color:theme.accent}});
          s.addShape(pres.shapes.RECTANGLE,{x:5,y:1.0,w:0.04,h:4.4,fill:{color:theme.accent+'44'},line:{color:theme.accent+'44'}});
          const half=Math.ceil((slide.bullets||[]).length/2);
          const col1=slide.bullets?.slice(0,half)||[];
          const col2=slide.bullets?.slice(half)||[];
          if(col1.length)s.addText(col1.map((b,j)=>({text:b,options:{bullet:true,fontSize:13,fontFace:"Calibri",color:theme.slideText,paraSpaceAfter:6,breakLine:j<col1.length-1}})),{x:0.5,y:1.1,w:4.3,h:4.2,valign:'top'});
          if(col2.length)s.addText(col2.map((b,j)=>({text:b,options:{bullet:true,fontSize:13,fontFace:"Calibri",color:theme.slideText,paraSpaceAfter:6,breakLine:j<col2.length-1}})),{x:5.2,y:1.1,w:4.3,h:4.2,valign:'top'});

        // ── BIG STAT ──
        } else if(tpl==='stat'){
          s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:10,h:0.06,fill:{color:theme.accent},line:{color:theme.accent}});
          await addLogo(s,pres,logoImg,logoPos||'top-left',logoWhiteBg,false);
          s.addText(slide.title,{x:0.5,y:0.15,w:9,h:0.7,fontSize:22,fontFace:"Calibri",bold:true,color:theme.slideText,align:"center",margin:0});
          const stat=(slide.bullets||[])[0]||'100%';
          s.addText(stat,{x:1,y:1.5,w:8,h:2.2,fontSize:72,fontFace:"Calibri",bold:true,color:theme.accent,align:'center'});
          const rest=(slide.bullets||[]).slice(1);
          if(rest.length)s.addText(rest.join('  ·  '),{x:1,y:3.9,w:8,h:0.8,fontSize:13,color:'666666',align:'center'});

        // ── DEFAULT (bullets + optional image) ──
        } else {
          s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:10,h:0.06,fill:{color:theme.accent},line:{color:theme.accent}});
          await addLogo(s,pres,logoImg,logoPos||'top-left',logoWhiteBg,false);
          if(imgData){
            s.addText(slide.title,{x:0.5,y:0.15,w:5.8,h:0.75,fontSize:20,fontFace:"Calibri",bold:true,color:theme.slideText,align:"left",margin:0});
            s.addShape(pres.shapes.RECTANGLE,{x:0.5,y:1.0,w:1.0,h:0.04,fill:{color:theme.accent},line:{color:theme.accent}});
            if(slide.bullets?.length)s.addText(slide.bullets.map((b,j)=>({text:b,options:{bullet:true,fontSize:13,fontFace:"Calibri",color:theme.slideText,paraSpaceAfter:6,breakLine:j<slide.bullets.length-1}})),{x:0.5,y:1.1,w:5.8,h:4.1,valign:'top'});
            try{s.addImage({data:imgData,x:6.4,y:0.06,w:3.6,h:5.565,sizing:{type:'cover',w:3.6,h:5.565}});}catch(e){}
          }else{
            s.addText(slide.title,{x:0.5,y:0.15,w:9,h:0.75,fontSize:24,fontFace:"Calibri",bold:true,color:theme.slideText,align:"left",margin:0});
            s.addShape(pres.shapes.RECTANGLE,{x:0.5,y:1.0,w:1.1,h:0.04,fill:{color:theme.accent},line:{color:theme.accent}});
            if(slide.bullets?.length)s.addText(slide.bullets.map((b,j)=>({text:b,options:{bullet:true,fontSize:14,fontFace:"Calibri",color:theme.slideText,paraSpaceAfter:7,breakLine:j<slide.bullets.length-1}})),{x:0.5,y:1.1,w:9,h:4.1,valign:'top'});
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
