export function createFireworks(canvas) {
  const ctx=canvas?.getContext('2d');
  let frame=0,particles=[],started=0,nextBurst=0,width=0,height=0,lastTime=0;
  const colors=['#f2c35e','#65e0bc','#87b7ff','#ed94a5','#c7a0ef'];
  const timings=[0,380,820,1300,1850,2450];
  const reduced=()=>window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  function stop(){cancelAnimationFrame(frame);frame=0;particles=[];if(canvas){canvas.hidden=true;ctx?.clearRect(0,0,width,height);}}
  function resize(){width=window.innerWidth;height=window.innerHeight;const dpr=Math.min(window.devicePixelRatio||1,2);canvas.width=Math.round(width*dpr);canvas.height=Math.round(height*dpr);canvas.style.width=width+'px';canvas.style.height=height+'px';ctx.setTransform(dpr,0,0,dpr,0,0);}
  function burst(){
    const x=width*(.18+Math.random()*.64),y=height*(.1+Math.random()*.34),color=colors[nextBurst%colors.length];
    for(let i=0;i<58;i++){const angle=Math.PI*2*i/58,speed=1.8+Math.random()*3.6;particles.push({x,y,vx:Math.cos(angle)*speed,vy:Math.sin(angle)*speed,life:1,decay:.015+Math.random()*.012,color});}
  }
  function draw(now){
    if(document.hidden||reduced()){stop();return;}
    const elapsed=now-started,dt=Math.min(2,(now-lastTime)/16.667||1);lastTime=now;
    while(nextBurst<timings.length&&elapsed>=timings[nextBurst]){burst();nextBurst++;}
    ctx.clearRect(0,0,width,height);
    for(const p of particles){const oldX=p.x,oldY=p.y;p.x+=p.vx*dt;p.y+=p.vy*dt;p.vy+=.065*dt;p.vx*=.99;p.life-=p.decay*dt;ctx.globalAlpha=Math.max(0,p.life);ctx.strokeStyle=p.color;ctx.lineWidth=2;ctx.beginPath();ctx.moveTo(oldX,oldY);ctx.lineTo(p.x,p.y);ctx.stroke();}
    ctx.globalAlpha=1;particles=particles.filter(p=>p.life>0);
    if(elapsed<4200&&(nextBurst<timings.length||particles.length))frame=requestAnimationFrame(draw);else stop();
  }
  function play(){stop();if(!ctx||reduced())return false;resize();canvas.hidden=false;started=performance.now();lastTime=started;nextBurst=0;frame=requestAnimationFrame(draw);return true;}
  window.addEventListener('resize',()=>{if(frame)resize();});
  document.addEventListener('visibilitychange',()=>{if(document.hidden)stop();});
  return {play,stop,get reducedMotion(){return reduced();}};
}
