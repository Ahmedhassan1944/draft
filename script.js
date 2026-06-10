/**
 * AI-Powered Automation Demo - Core Logic
 */

document.addEventListener('DOMContentLoaded', () => {
    initParticles();
    initScrollProgress();
    initScrollReveal();
    initStoryAnimations();
});

/* --- SCROLL PROGRESS --- */
function initScrollProgress() {
    const progressBar = document.getElementById('scroll-progress');
    window.addEventListener('scroll', () => {
        const winScroll = document.body.scrollTop || document.documentElement.scrollTop;
        const height = document.documentElement.scrollHeight - document.documentElement.clientHeight;
        const scrolled = (winScroll / height) * 100;
        progressBar.style.width = scrolled + '%';
    });
}

/* --- SCROLL REVEAL (Intersection Observer) --- */
function initScrollReveal() {
    const reveals = document.querySelectorAll('.reveal');
    
    const observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.classList.add('active');
            }
        });
    }, {
        threshold: 0.15
    });

    reveals.forEach(reveal => {
        observer.observe(reveal);
    });
}

/* --- STORY ANIMATIONS --- */
function initStoryAnimations() {
    // Flags to ensure animations only run once when triggered
    let step1Triggered = false;
    let step2Triggered = false;
    let step4Triggered = false;
    let step6Triggered = false;
    let step7Triggered = false;

    const observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                const id = entry.target.id;
                
                if (id === 'step-1-pain' && !step1Triggered) {
                    step1Triggered = true;
                    runPainAnimation();
                } else if (id === 'step-2-translation' && !step2Triggered) {
                    step2Triggered = true;
                    runTranslationTyping();
                } else if (id === 'step-4-5-generation' && !step4Triggered) {
                    step4Triggered = true;
                    runCodeGeneration();
                } else if (id === 'step-6-execution' && !step6Triggered) {
                    step6Triggered = true;
                    runExecutionAnimation();
                } else if (id === 'step-7-roi' && !step7Triggered) {
                    step7Triggered = true;
                    runROIAnimation();
                }
            }
        });
    }, { threshold: 0.5 });

    const sections = document.querySelectorAll('.story-section');
    sections.forEach(sec => observer.observe(sec));
}

/* --- SPECIFIC SECTION ANIMATIONS --- */

// Step 1: The Pain (Looping)
function runPainAnimation() {
    const idEl = document.querySelector('.typing-id');
    const counterEl = document.getElementById('pain-counter-val');
    let count = 400;
    
    setInterval(() => {
        const randomId = Math.floor(10000 + Math.random() * 90000);
        idEl.textContent = randomId;
        
        count--;
        if(count < 0) count = 400;
        counterEl.textContent = count;
    }, 1500); // Slow, painful loop
}

// Step 2: Translation (Typing)
function runTranslationTyping() {
    const container = document.getElementById('typing-container');
    const lines = [
        "I open the Employee_Data_Final_v4.xlsx workbook.",
        "I find the next empty employee row.",
        "I copy their ID number.",
        "I paste it into the reporting template.",
        "I wait for the VLOOKUPs to fetch the data.",
        "I press print to generate the PDF form.",
        "I repeat this process 400 times every month."
    ];
    
    let lineIdx = 0;
    
    function typeLine() {
        if (lineIdx >= lines.length) {
            // Done typing, add pulsing cursor
            container.innerHTML += '<span class="typing-cursor"></span>';
            return;
        }
        
        const line = lines[lineIdx];
        const lineEl = document.createElement('div');
        container.appendChild(lineEl);
        
        let charIdx = 0;
        
        const typeChar = setInterval(() => {
            if (charIdx >= line.length) {
                clearInterval(typeChar);
                lineIdx++;
                setTimeout(typeLine, 500); // Pause between lines
            } else {
                lineEl.textContent += line[charIdx];
                charIdx++;
            }
        }, 30); // Typing speed
    }
    
    typeLine();
}

// Step 4/5: Code Generation
function runCodeGeneration() {
    const container = document.getElementById('code-generation-container');
    const vbaCode = `
<span class="token comment">' ============================================================</span>
<span class="token comment">' Main Macro: Auto-print employee forms from a range</span>
<span class="token comment">' ============================================================</span>
<span class="token keyword">Sub</span> <span class="token function">PrintEmployeeRange</span>()

    <span class="token comment">' ---- Variable declarations ----</span>
    <span class="token keyword">Dim</span> ws           <span class="token keyword">As</span> Worksheet
    <span class="token keyword">Dim</span> firstEmp     <span class="token keyword">As Long</span>
    <span class="token keyword">Dim</span> lastEmp      <span class="token keyword">As Long</span>
    <span class="token keyword">Dim</span> empNum       <span class="token keyword">As Long</span>
    <span class="token keyword">Dim</span> totalPrinted <span class="token keyword">As Long</span>
    <span class="token keyword">Dim</span> userInput    <span class="token keyword">As String</span>
    <span class="token keyword">Dim</span> confirmMsg   <span class="token keyword">As String</span>

    <span class="token comment">' ---- Error handling ----</span>
    <span class="token keyword">On Error GoTo</span> ErrorHandler

    <span class="token comment">' ---- Reference the form sheet ----</span>
    <span class="token keyword">Set</span> ws = ThisWorkbook.Worksheets(<span class="token string">"x"</span>)

    <span class="token comment">' ==== Step 1: Get first employee number ====</span>
    userInput = <span class="token function">InputBox</span>( _
        <span class="token string">"Enter the FIRST employee number:"</span> & vbNewLine & _
        <span class="token string">"(Example: 101)"</span>, _
        <span class="token string">"Print Employee Forms - Step 1/2"</span>)

    <span class="token keyword">If</span> userInput = <span class="token string">""</span> <span class="token keyword">Then</span>
        <span class="token function">MsgBox</span> <span class="token string">"Operation cancelled."</span>, vbInformation, <span class="token string">"Cancelled"</span>
        <span class="token keyword">Exit Sub</span>
    <span class="token keyword">End If</span>

    <span class="token keyword">If Not</span> <span class="token function">IsNumeric</span>(userInput) <span class="token keyword">Then</span>
        <span class="token function">MsgBox</span> <span class="token string">"Error: Please enter a valid number."</span>, vbCritical, <span class="token string">"Input Error"</span>
        <span class="token keyword">Exit Sub</span>
    <span class="token keyword">End If</span>
    firstEmp = <span class="token function">CLng</span>(userInput)

    <span class="token comment">' ==== Step 2: Get last employee number ====</span>
    userInput = <span class="token function">InputBox</span>( _
        <span class="token string">"Enter the LAST employee number:"</span> & vbNewLine & _
        <span class="token string">"(Example: 115)"</span>, _
        <span class="token string">"Print Employee Forms - Step 2/2"</span>)

    <span class="token keyword">If</span> userInput = <span class="token string">""</span> <span class="token keyword">Then</span>
        <span class="token function">MsgBox</span> <span class="token string">"Operation cancelled."</span>, vbInformation, <span class="token string">"Cancelled"</span>
        <span class="token keyword">Exit Sub</span>
    <span class="token keyword">End If</span>

    <span class="token keyword">If Not</span> <span class="token function">IsNumeric</span>(userInput) <span class="token keyword">Then</span>
        <span class="token function">MsgBox</span> <span class="token string">"Error: Please enter a valid number."</span>, vbCritical, <span class="token string">"Input Error"</span>
        <span class="token keyword">Exit Sub</span>
    <span class="token keyword">End If</span>
    lastEmp = <span class="token function">CLng</span>(userInput)

    <span class="token comment">' ---- Validate range ----</span>
    <span class="token keyword">If</span> firstEmp > lastEmp <span class="token keyword">Then</span>
        <span class="token function">MsgBox</span> <span class="token string">"Error: First number must be less than or equal to last number."</span>, _
               vbCritical, <span class="token string">"Input Error"</span>
        <span class="token keyword">Exit Sub</span>
    <span class="token keyword">End If</span>

    <span class="token comment">' ==== Step 3: Confirm before printing ====</span>
    confirmMsg = <span class="token string">"Forms will be printed from employee "</span> & firstEmp & _
                 <span class="token string">" to employee "</span> & lastEmp & vbNewLine & _
                 <span class="token string">"Total forms: "</span> & (lastEmp - firstEmp + 1) & vbNewLine & vbNewLine & _
                 <span class="token string">"Do you want to continue?"</span>

    <span class="token keyword">If</span> <span class="token function">MsgBox</span>(confirmMsg, vbQuestion + vbYesNo, <span class="token string">"Confirm Print"</span>) = vbNo <span class="token keyword">Then</span>
        <span class="token function">MsgBox</span> <span class="token string">"Operation cancelled."</span>, vbInformation, <span class="token string">"Cancelled"</span>
        <span class="token keyword">Exit Sub</span>
    <span class="token keyword">End If</span>

    <span class="token comment">' ---- Disable screen updating to speed up the process ----</span>
    Application.ScreenUpdating = <span class="token keyword">False</span>
    Application.Calculation = xlCalculationAutomatic

    totalPrinted = 0

    <span class="token comment">' ==== Step 4: Loop through each employee number ====</span>
    <span class="token keyword">For</span> empNum = firstEmp <span class="token keyword">To</span> lastEmp

        <span class="token comment">' -- Set employee number in cell F21 --</span>
        ws.Range(<span class="token string">"F21"</span>).Value = empNum

        <span class="token comment">' -- Force Excel to recalculate XLOOKUP formulas --</span>
        Application.Calculate

        <span class="token comment">' -- Print the form sheet --</span>
        ws.PrintOut Copies:=1, Collate:=<span class="token keyword">True</span>

        <span class="token comment">' -- Increment printed counter --</span>
        totalPrinted = totalPrinted + 1

    <span class="token keyword">Next</span> empNum

    <span class="token comment">' ---- Re-enable screen updating ----</span>
    Application.ScreenUpdating = <span class="token keyword">True</span>

    <span class="token comment">' ==== Step 5: Success message ====</span>
    <span class="token function">MsgBox</span> <span class="token string">"Printing completed successfully!"</span> & vbNewLine & vbNewLine & _
           <span class="token string">"Total forms printed: "</span> & totalPrinted & vbNewLine & _
           <span class="token string">"(Employee "</span> & firstEmp & <span class="token string">" to "</span> & lastEmp & <span class="token string">")"</span>, _
           vbInformation, <span class="token string">"Done"</span>

    <span class="token keyword">Set</span> ws = <span class="token keyword">Nothing</span>
    <span class="token keyword">Exit Sub</span>

<span class="token comment">' ==== Error Handler ====</span>
ErrorHandler:
    Application.ScreenUpdating = <span class="token keyword">True</span>
    <span class="token function">MsgBox</span> <span class="token string">"An unexpected error occurred!"</span> & vbNewLine & _
           <span class="token string">"Error number: "</span> & Err.Number & vbNewLine & _
           <span class="token string">"Description: "</span> & Err.Description, _
           vbCritical, <span class="token string">"Error"</span>
    <span class="token keyword">Set</span> ws = <span class="token keyword">Nothing</span>

<span class="token keyword">End Sub</span>
`;
    // Split by tags and text to simulate typing without breaking HTML
    // A simple trick: inject full HTML, but hide it and reveal character by character
    container.innerHTML = vbaCode;
    const content = container.innerHTML;
    container.innerHTML = '';
    
    let i = 0;
    // We will just type text, but for syntax highlighting we'd need a more complex parser.
    // For this demo, we'll slice the HTML string directly and rely on browser fixing open tags,
    // or better, just reveal it chunk by chunk rapidly.
    
    const chunkTyping = setInterval(() => {
        if (i >= content.length) {
            clearInterval(chunkTyping);
        } else {
            // take next 5 chars for speed
            const end = Math.min(i + 15, content.length);
            container.innerHTML = content.substring(0, end);
            i += 15;
        }
    }, 20); // Extremely fast generation
}

// Step 6: Execution Loop
function runExecutionAnimation() {
    const countEl = document.getElementById('forms-processed-count');
    const percEl = document.getElementById('forms-percentage');
    const barEl = document.getElementById('execution-progress');
    const logsEl = document.getElementById('execution-logs');
    
    let processed = 0;
    const total = 400;
    
    const executionInterval = setInterval(() => {
        if (processed >= total) {
            clearInterval(executionInterval);
            logsEl.innerHTML += `<div class="log-line text-gold">>>> AUTOMATION COMPLETE. TIME ELAPSED: 2.1s</div>`;
            return;
        }
        
        processed += Math.floor(Math.random() * 5) + 1; // Process 1-5 at a time
        if (processed > total) processed = total;
        
        const percentage = Math.floor((processed / total) * 100);
        
        countEl.textContent = processed;
        percEl.textContent = percentage + '%';
        barEl.style.width = percentage + '%';
        
        // Add log
        const id = Math.floor(10000 + Math.random() * 90000);
        const ms = Math.floor(Math.random() * 10) + 1;
        const log = document.createElement('div');
        log.className = 'log-line';
        log.textContent = `[SUCCESS] Processed ID ${id} -> PDF Exported (${ms}ms)`;
        logsEl.appendChild(log);
        
        // Auto scroll logs
        if (logsEl.children.length > 6) {
            logsEl.removeChild(logsEl.firstChild);
        }
        
    }, 50); // Very fast execution
}

// Step 7: ROI Counter
function runROIAnimation() {
    const counters = document.querySelectorAll('.count-up');
    const duration = 2000; // 2 seconds
    
    counters.forEach(counter => {
        const target = +counter.getAttribute('data-target');
        const increment = target / (duration / 16); // 60fps
        
        let current = 0;
        
        const updateCounter = () => {
            current += increment;
            if (current < target) {
                counter.innerText = Math.ceil(current).toLocaleString();
                requestAnimationFrame(updateCounter);
            } else {
                counter.innerText = target.toLocaleString();
            }
        };
        
        updateCounter();
    });
}

/* --- CANVAS PARTICLES --- */
function initParticles() {
    const canvas = document.getElementById('bg-canvas');
    const ctx = canvas.getContext('2d');
    
    let width, height;
    let particles = [];
    
    function resize() {
        width = window.innerWidth;
        height = window.innerHeight;
        canvas.width = width;
        canvas.height = height;
    }
    
    window.addEventListener('resize', resize);
    resize();
    
    class Particle {
        constructor() {
            this.x = Math.random() * width;
            this.y = Math.random() * height;
            this.size = Math.random() * 2 + 0.5;
            this.speedX = (Math.random() - 0.5) * 0.5;
            this.speedY = (Math.random() - 0.5) * 0.5;
            this.opacity = Math.random() * 0.5 + 0.1;
            
            // Random color from palette
            const colors = ['#3b82f6', '#8b5cf6', '#10b981'];
            this.color = colors[Math.floor(Math.random() * colors.length)];
        }
        
        update() {
            this.x += this.speedX;
            this.y += this.speedY;
            
            // Wrap around
            if (this.x > width) this.x = 0;
            if (this.x < 0) this.x = width;
            if (this.y > height) this.y = 0;
            if (this.y < 0) this.y = height;
        }
        
        draw() {
            ctx.beginPath();
            ctx.arc(this.x, this.y, this.size, 0, Math.PI * 2);
            ctx.fillStyle = this.color;
            ctx.globalAlpha = this.opacity;
            ctx.fill();
        }
    }
    
    function createParticles() {
        // Adjust amount based on screen size for performance
        const amount = Math.floor((width * height) / 15000);
        for (let i = 0; i < amount; i++) {
            particles.push(new Particle());
        }
    }
    
    function animate() {
        ctx.clearRect(0, 0, width, height);
        
        for (let i = 0; i < particles.length; i++) {
            particles[i].update();
            particles[i].draw();
        }
        
        // Connecting lines for AI vibe
        ctx.globalAlpha = 0.05;
        ctx.strokeStyle = '#ffffff';
        ctx.lineWidth = 0.5;
        
        for (let i = 0; i < particles.length; i++) {
            for (let j = i; j < particles.length; j++) {
                const dx = particles[i].x - particles[j].x;
                const dy = particles[i].y - particles[j].y;
                const distance = Math.sqrt(dx * dx + dy * dy);
                
                if (distance < 100) {
                    ctx.beginPath();
                    ctx.moveTo(particles[i].x, particles[i].y);
                    ctx.lineTo(particles[j].x, particles[j].y);
                    ctx.stroke();
                }
            }
        }
        
        requestAnimationFrame(animate);
    }
    
    createParticles();
    animate();
}
