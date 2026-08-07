document.addEventListener('DOMContentLoaded', async function() {
    const themeToggle = document.getElementById('checkbox');
    const body = document.body;
    
    if (localStorage.getItem('theme') === 'dark') {
        body.classList.remove('theme-light');
        body.classList.add('theme-dark');
        themeToggle.checked = true;
    }
    
    themeToggle.addEventListener('change', function() {
        if(this.checked) {
            body.classList.replace('theme-light', 'theme-dark');
            localStorage.setItem('theme', 'dark');
        } else {
            body.classList.replace('theme-dark', 'theme-light');
            localStorage.setItem('theme', 'light');
        }
    });

    function showToast(message, type = 'success') {
        const toastContainer = document.querySelector('.toast-container');
        const toast = document.createElement('div');
        toast.className = `toast align-items-center text-white bg-${type} border-0 show`;
        toast.setAttribute('role', 'alert');
        
        toast.innerHTML = `
            <div class="d-flex">
                <div class="toast-body">${message}</div>
                <button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast"></button>
            </div>
        `;
        toastContainer.appendChild(toast);
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => toast.remove(), 300);
        }, 3000);
    }

    const enableInstructorCheckbox = document.getElementById('enable_instructor');
    const instructorDetails = document.getElementById('instructor-details');
    const instructorNameInput = document.getElementById('instructor_name');
    const instructorTitleInput = document.getElementById('instructor_title');

    enableInstructorCheckbox.addEventListener('change', function() {
        if (this.checked) {
            instructorDetails.style.display = 'block';
            void instructorDetails.offsetWidth; 
            instructorDetails.style.opacity = '1';
            instructorNameInput.setAttribute('required', 'required');
            instructorTitleInput.setAttribute('required', 'required');
        } else {
            instructorDetails.style.opacity = '0';
            setTimeout(() => {
                instructorDetails.style.display = 'none';
                instructorNameInput.removeAttribute('required');
                instructorTitleInput.removeAttribute('required');
            }, 400);
        }
    });

    const templateSelector = document.getElementById('template_type');
    const ministryFields = document.querySelectorAll('.template-ministry-only');
    
    if(templateSelector) {
        templateSelector.addEventListener('change', function() {
            if (this.value === 'ministry') {
                ministryFields.forEach(el => el.style.display = 'block');
            } else {
                ministryFields.forEach(el => el.style.display = 'none');
            }
        });
    }

    const dropExcel = document.getElementById('drop-excel');
    const excelInput = document.getElementById('excel_file');
    const excelFileName = dropExcel.querySelector('.file-name');
    
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropExcel.addEventListener(eventName, e => { e.preventDefault(); e.stopPropagation(); });
    });

    ['dragenter', 'dragover'].forEach(eventName => {
        dropExcel.addEventListener(eventName, () => dropExcel.classList.add('dragover'));
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropExcel.addEventListener(eventName, () => dropExcel.classList.remove('dragover'));
    });

    dropExcel.addEventListener('drop', (e) => {
        excelInput.files = e.dataTransfer.files;
        if (excelInput.files.length > 0) {
            excelFileName.textContent = excelInput.files[0].name;
        }
    });

    excelInput.addEventListener('change', () => {
        if (excelInput.files.length > 0) {
            excelFileName.textContent = excelInput.files[0].name;
        }
    });

    const instLogoInput = document.getElementById('instructor_logo');
    instLogoInput.addEventListener('change', () => {
        if (instLogoInput.files.length > 0) showToast('تم اختيار لوجو المدرب', 'info');
    });

    // --- LUXURIOUS CERTIFICATE GENERATION LOGIC ---
    const CANVAS_WIDTH = 3508;
    const CANVAS_HEIGHT = 2480;
    
    // Luxurious Navy & Gold Colors
    const COLOR_GOLD = "rgb(184, 134, 11)"; // Darker Gold
    const COLOR_NAVY = "rgb(10, 25, 49)"; // Dark Navy for texts
    const COLOR_WHITE = "rgb(255, 255, 255)";
    const COLOR_LIGHT_GOLD = "rgb(184, 134, 11)";
    
    const CEO_NAME = "Ammar Murad";

    async function loadFonts() {
        await document.fonts.load('240px "Cinzel"');
        await document.fonts.load('140px "Great Vibes"');
        await document.fonts.load('100px "Tinos"');
    }

    function loadLocalImage(path) {
        return new Promise((resolve) => {
            const img = new Image();
            img.onload = () => resolve(img);
            img.onerror = () => resolve(null); 
            img.src = path;
        });
    }

    function removeWhiteBackground(img) {
        if (!img) return null;
        const c = document.createElement('canvas');
        c.width = img.width;
        c.height = img.height;
        const x = c.getContext('2d');
        x.drawImage(img, 0, 0);
        const imgData = x.getImageData(0, 0, c.width, c.height);
        const data = imgData.data;
        for (let i = 0; i < data.length; i += 4) {
            let r = data[i], g = data[i+1], b = data[i+2];
            // If pixel is very close to white, make it transparent
            if (r > 230 && g > 230 && b > 230) {
                data[i+3] = 0; // alpha = 0
            }
        }
        x.putImageData(imgData, 0, 0);
        const newImg = new Image();
        return new Promise((resolve) => {
            newImg.onload = () => resolve(newImg);
            newImg.src = c.toDataURL("image/png");
        });
    }

    function readUploadedImage(file) {
        return new Promise((resolve) => {
            if (!file) return resolve(null);
            const reader = new FileReader();
            reader.onload = (e) => {
                const img = new Image();
                img.onload = () => resolve(img);
                img.src = e.target.result;
            };
            reader.readAsDataURL(file);
        });
    }

    function readExcel(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, {type: 'array'});
                    const firstSheet = workbook.SheetNames[0];
                    const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], {defval: ""});
                    resolve(jsonData);
                } catch (err) { reject(err); }
            };
            reader.readAsArrayBuffer(file);
        });
    }

    function drawCenteredText(ctx, text, y_pos, fontString, color, canvas_width = CANVAS_WIDTH) {
        ctx.font = fontString;
        ctx.fillStyle = color;
        const metrics = ctx.measureText(text);
        const x_pos = (canvas_width - metrics.width) / 2;
        ctx.fillText(text, x_pos, y_pos);
        const fontSize = parseInt(fontString.match(/\d+px/)[0]);
        return y_pos + fontSize + (fontSize * 0.2); 
    }

    function drawWrappedText(ctx, text, y_pos, fontString, color, max_width, canvas_width = CANVAS_WIDTH, line_height_mult = 1.3) {
        ctx.font = fontString;
        ctx.fillStyle = color;
        const words = text.split(" ");
        let lines = [];
        let current_line = words[0];

        for (let i = 1; i < words.length; i++) {
            let word = words[i];
            let width = ctx.measureText(current_line + " " + word).width;
            if (width <= max_width) {
                current_line += " " + word;
            } else {
                lines.push(current_line);
                current_line = word;
            }
        }
        lines.push(current_line);

        let curr_y = y_pos;
        const fontSize = parseInt(fontString.match(/\d+px/)[0]);
        for (let line of lines) {
            drawCenteredText(ctx, line, curr_y, fontString, color, canvas_width);
            curr_y += fontSize * line_height_mult; 
        }
        return curr_y;
    }

    function drawCenteredUnderLine(ctx, text, line_start_x, line_length, y_pos, fontString, color) {
        ctx.font = fontString;
        ctx.fillStyle = color;
        const text_w = ctx.measureText(text).width;
        const x_pos = line_start_x + (line_length - text_w) / 2;
        ctx.fillText(text, x_pos, y_pos);
        const fontSize = parseInt(fontString.match(/\d+px/)[0]);
        return y_pos + fontSize + (fontSize * 0.2);
    }
    
    function scaleImage(img, maxW, maxH) {
        let ratio = Math.min(maxW / img.width, maxH / img.height);
        return { w: img.width * ratio, h: img.height * ratio };
    }

    function getFittedFont(ctx, text, max_width, initial_size, min_size, fontFamily) {
        let size = initial_size;
        while (size >= min_size) {
            ctx.font = `${size}px "${fontFamily}"`;
            if (ctx.measureText(text).width <= max_width) return `${size}px "${fontFamily}"`;
            size -= 4;
        }
        return `${min_size}px "${fontFamily}"`;
    }

    const form = document.getElementById('generator-form');
    
    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        
        const courseName = document.getElementById('course_name').value.trim();
        const issueMonth = document.getElementById('issue_month').value;
        const issueYear = document.getElementById('issue_year').value;
        const templateType = document.getElementById('template_type') ? document.getElementById('template_type').value : 'scientific_center';
        
        const subjectName = document.getElementById('subject_name') ? document.getElementById('subject_name').value.trim() : '';
        const startDate = document.getElementById('start_date') ? document.getElementById('start_date').value : '';
        const endDate = document.getElementById('end_date') ? document.getElementById('end_date').value : '';
        const grade = document.getElementById('grade') ? document.getElementById('grade').value : '';
        const certificateId = document.getElementById('certificate_id') ? document.getElementById('certificate_id').value.trim() : '';
        const studentTitle = document.getElementById('student_title') ? document.getElementById('student_title').value : 'Mr./Ms. ';
        
        if (!courseName) { showToast('الرجاء كتابة اسم الدورة', 'danger'); return; }

        let studentsList = [];
        const activeTab = document.querySelector('#studentTabs .active').id;
        
        if (activeTab === 'manual-tab') {
            const singleName = document.getElementById('single_student_name').value.trim();
            if (!singleName) { showToast('الرجاء كتابة اسم الطالب', 'danger'); return; }
            studentsList.push({ name: singleName });
        } else {
            const excelFile = document.getElementById('excel_file').files[0];
            if (!excelFile) { showToast('الرجاء رفع ملف الإكسيل', 'danger'); return; }
            try {
                const df = await readExcel(excelFile);
                if (df.length === 0) throw new Error("الملف فارغ");
                
                const keys = Object.keys(df[0]);
                let nameKey = null;
                for (let p of ["Student Name", "Name", "Student", "الاسم", "اسم الطالب"]) {
                    let lowerP = p.toLowerCase();
                    for (let k of keys) {
                        if (k.toLowerCase().includes(lowerP)) { nameKey = k; break; }
                    }
                    if (nameKey) break;
                }
                
                if (!nameKey) throw new Error("لم يتم العثور على عمود الأسماء في الإكسيل");
                
                df.forEach(row => {
                    let n = (row[nameKey] || "").toString().trim();
                    if (n && n.toLowerCase() !== 'nan') studentsList.push({ name: n });
                });
            } catch (err) {
                showToast(`خطأ في قراءة الإكسيل: ${err.message}`, 'danger');
                return;
            }
        }

        if (studentsList.length === 0) { showToast('لم يتم العثور على أي طلاب', 'danger'); return; }

        const useInstructor = document.getElementById('enable_instructor').checked;
        const instName = document.getElementById('instructor_name').value.trim();
        const instTitle = document.getElementById('instructor_title').value.trim();
        const instLogoFile = document.getElementById('instructor_logo').files[0];

        const submitBtn = document.getElementById('process-btn');
        submitBtn.disabled = true;
        
        const progressContainer = document.getElementById('progress-container');
        const mainDashboard = document.getElementById('main-dashboard');
        const headerSection = document.querySelector('.header-section');
        const resultContainer = document.getElementById('result-container');
        
        mainDashboard.classList.add('d-none');
        headerSection.classList.add('d-none');
        progressContainer.classList.remove('d-none');
        
        const progressBar = document.getElementById('progress-bar');
        const progressStatus = document.getElementById('progress-status');
        const currentStudentDiv = document.getElementById('current-student');

        try {
            progressStatus.textContent = "جاري تحميل الخطوط الفخمة...";
            await loadFonts();

            progressStatus.textContent = "جاري قراءة الصور الثابتة...";
            const bacImg = await loadLocalImage('bac.png');
            let mainLogoImg = await loadLocalImage('logo.jpeg');
            let depImg = await loadLocalImage('dep.png');
            let ceoLogoImg = await loadLocalImage('logo1.png');
            let ministryLogoImg = await loadLocalImage('sc.png');
            let instLogoImg = null;
            if (useInstructor && instLogoFile) {
                instLogoImg = await readUploadedImage(instLogoFile);
            }

            progressStatus.textContent = "جاري معالجة خلفيات الصور...";
            mainLogoImg = await removeWhiteBackground(mainLogoImg);
            depImg = await removeWhiteBackground(depImg);
            ceoLogoImg = await removeWhiteBackground(ceoLogoImg);
            ministryLogoImg = await removeWhiteBackground(ministryLogoImg);
            if (instLogoImg) instLogoImg = await removeWhiteBackground(instLogoImg);

            const doc = new jspdf.jsPDF({
                orientation: 'landscape',
                unit: 'pt', 
                format: [CANVAS_WIDTH, CANVAS_HEIGHT] 
            });

            const totalStudents = studentsList.length;
            const startTime = Date.now();

            const canvas = document.createElement('canvas');
            canvas.width = CANVAS_WIDTH;
            canvas.height = CANVAS_HEIGHT;
            const ctx = canvas.getContext('2d');
            
            const delay = ms => new Promise(res => setTimeout(res, ms));

            for (let i = 0; i < totalStudents; i++) {
                let studentName = studentsList[i].name;
                studentName = studentName.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');

                let percent = Math.round(((i + 1) / totalStudents) * 100);
                progressBar.style.width = `${percent}%`;
                progressBar.textContent = `${percent}%`;
                currentStudentDiv.innerHTML = `<i class="fas fa-user-graduate me-2"></i>${studentName} (${i + 1}/${totalStudents})`;
                progressStatus.textContent = "جاري تصميم الشهادة...";
                
                await delay(10); 

                // --- LUXURIOUS DRAWING START --- //

                if (templateType === 'ministry') {
                    // 1. Background
                    ctx.fillStyle = "#ffffff";
                    ctx.fillRect(0, 0, CANVAS_WIDTH, CANVAS_HEIGHT);
                    
                    // 2. Borders
                    ctx.strokeStyle = COLOR_GOLD;
                    ctx.lineWidth = 15;
                    ctx.strokeRect(80, 80, CANVAS_WIDTH - 160, CANVAS_HEIGHT - 160);
                    ctx.lineWidth = 5;
                    ctx.strokeRect(110, 110, CANVAS_WIDTH - 220, CANVAS_HEIGHT - 220);
                    
                    // 3. Watermark
                    if (ministryLogoImg) {
                        ctx.globalAlpha = 0.20; // Increased opacity
                        let s = scaleImage(ministryLogoImg, 1400, 1400);
                        ctx.drawImage(ministryLogoImg, (CANVAS_WIDTH - s.w) / 2, (CANVAS_HEIGHT - s.h) / 2 + 50, s.w, s.h);
                        ctx.globalAlpha = 1.0;
                    }

                    // 4. Top Badges and Info
                    // Logo on Top Left
                    if (ministryLogoImg) {
                        let s = scaleImage(ministryLogoImg, 550, 550); // Increased size
                        ctx.drawImage(ministryLogoImg, 150, 150, s.w, s.h);
                    }

                    // Institution Info on Top Right
                    let instInfoY = 150;
                    let instInfoX = CANVAS_WIDTH - 650;
                    ctx.textAlign = "center";
                    ctx.textBaseline = "top";
                    
                    // 1. Ministry of Education
                    ctx.font = 'bold 45px Arial, sans-serif';
                    ctx.fillStyle = "#1F355E"; 
                    ctx.fillText("Ministry of Education", instInfoX, instInfoY);
                    
                    // 2. Directorate...
                    ctx.fillText("Directorate of Education in Beni Suef", instInfoX, instInfoY + 60);
                    
                    // 3. Private Education
                    ctx.fillStyle = "#8B1E24"; 
                    ctx.fillText("Private Education", instInfoX, instInfoY + 120);
                    
                    // Separator Line
                    ctx.strokeStyle = "#000000";
                    ctx.lineWidth = 4;
                    ctx.beginPath();
                    ctx.moveTo(instInfoX - 300, instInfoY + 190);
                    ctx.lineTo(instInfoX + 300, instInfoY + 190);
                    ctx.stroke();
                    
                    // 4. Scientific Center...
                    ctx.fillStyle = "#1F355E"; 
                    ctx.fillText("Scientific Center for Computer & Languages", instInfoX, instInfoY + 220);
                    
                    // 5. Approved by...
                    ctx.fillStyle = "#8B1E24"; 
                    ctx.fillText("Approved by Ministry of Education", instInfoX, instInfoY + 280);
                    
                    ctx.textAlign = "left";

                    // 5. Title
                    let curr_y = 480; 
                    curr_y = drawCenteredText(ctx, "CERTIFICATE", curr_y, '180px "Cinzel"', "#8B1E24") + 20;
                    curr_y = drawCenteredText(ctx, "OF", curr_y, '70px "Cinzel"', "#8B1E24") + 10;
                    curr_y = drawCenteredText(ctx, "COMPLETION", curr_y, '130px "Cinzel"', "#8B1E24") + 50;
                    
                    let gradLine = ctx.createLinearGradient(CANVAS_WIDTH/2 - 500, 0, CANVAS_WIDTH/2 + 500, 0);
                    gradLine.addColorStop(0, "rgba(139, 30, 36, 0)");
                    gradLine.addColorStop(0.5, "#8B1E24");
                    gradLine.addColorStop(1, "rgba(139, 30, 36, 0)");
                    ctx.fillStyle = gradLine;
                    ctx.fillRect(CANVAS_WIDTH/2 - 500, curr_y, 1000, 3);
                    curr_y += 80;

                    // 6. Recipient
                    ctx.font = '65px "Tinos"';
                    let nameTitleW = ctx.measureText(studentTitle).width;
                    ctx.font = 'bold 90px "Tinos"';
                    let nameW = ctx.measureText(studentName).width;
                    let totalNameW = nameTitleW + nameW;
                    let startNameX = (CANVAS_WIDTH - totalNameW) / 2;
                    
                    ctx.font = '65px "Tinos"';
                    ctx.fillStyle = "#1F355E";
                    ctx.fillText(studentTitle, startNameX, curr_y);
                    ctx.font = 'bold 90px "Tinos"';
                    ctx.fillStyle = "#8B1E24";
                    ctx.fillText(studentName, startNameX + nameTitleW, curr_y);
                    curr_y += 130;

                    // 7. Course Section
                    curr_y = drawCenteredText(ctx, `Successfully completed ${courseName}`, curr_y, '55px "Tinos"', "#1F355E") + 20;
                    curr_y = drawCenteredText(ctx, `${subjectName}`, curr_y, 'bold 75px "Tinos"', "#8B1E24") + 80;

                    // 8. Appreciation Paragraph
                    let appreciationText = "This certificate is proudly presented in recognition of the successful completion of the training program. We sincerely appreciate your dedication, perseverance, and commitment throughout the course. Congratulations on this achievement, and we wish you continued success and excellence in your future endeavors.";
                    curr_y = drawWrappedText(ctx, appreciationText, curr_y, 'italic 55px "Tinos"', "#1F355E", CANVAS_WIDTH - 600, CANVAS_WIDTH, 1.4) + 60;

                    // 9. Duration & Grade
                    curr_y = drawCenteredText(ctx, `Duration: ${startDate} to ${endDate}   |   Final Grade: ${grade}%`, curr_y, 'bold 50px "Tinos"', "#8B1E24") + 40;

                    // 10. Certificate ID
                    curr_y = drawCenteredText(ctx, `ID No.: ${certificateId}`, curr_y, '45px "Outfit", sans-serif', "#1F355E") + 150;

                    // 11. Signatures
                    let sigY = CANVAS_HEIGHT - 320;
                    ctx.strokeStyle = "#1F355E";
                    ctx.lineWidth = 2;
                    
                    let execX = 600;
                    ctx.beginPath(); ctx.moveTo(execX - 300, sigY); ctx.lineTo(execX + 300, sigY); ctx.stroke();
                    drawCenteredUnderLine(ctx, "Executive Director", execX - 300, 600, sigY + 20, '50px "Tinos"', "#1F355E");

                    let dirX = CANVAS_WIDTH / 2;
                    ctx.beginPath(); ctx.moveTo(dirX - 300, sigY); ctx.lineTo(dirX + 300, sigY); ctx.stroke();
                    drawCenteredUnderLine(ctx, "Director of Private Education", dirX - 300, 600, sigY + 20, '50px "Tinos"', "#1F355E");

                    let genX = CANVAS_WIDTH - 600;
                    ctx.beginPath(); ctx.moveTo(genX - 300, sigY); ctx.lineTo(genX + 300, sigY); ctx.stroke();
                    drawCenteredUnderLine(ctx, "General Director", genX - 300, 600, sigY + 20, '50px "Tinos"', "#1F355E");
                } else {
                    // 1. Background Image
                    if (bacImg) {
                        ctx.drawImage(bacImg, 0, 0, CANVAS_WIDTH, CANVAS_HEIGHT);
                    } else {
                        ctx.fillStyle = "rgb(255, 255, 255)"; // Fallback
                        ctx.fillRect(0, 0, CANVAS_WIDTH, CANVAS_HEIGHT);
                    }

                    const mInner = 150; 

                    // 2. Top Logos
                    // Top Right: Accreditations (dep.png)
                    // Top Left: Accreditations (dep.png)
    if (depImg) {
        let s = scaleImage(mainLogoImg, 700, 450);

        let x = CANVAS_WIDTH - s.w - (mInner - 50);
        let y = mInner - 50;

        ctx.drawImage(mainLogoImg, x, y, s.w, s.h);
    }

    // Top Right Logo
    if (mainLogoImg) {
          let s = scaleImage(depImg, 1600, 650);

        // نفس الـ Y الخاص باللوجو الثاني
        let x = mInner - 75;
        let y = mInner - 50;

        ctx.drawImage(depImg, x, y, s.w, s.h);
    }

                    // 3. Center Content & Typography
                    let curr_y = mInner + 140; 
                    ctx.textBaseline = "top";
                    
                    // Add scientific center above CERTIFICATE
                    curr_y = drawCenteredText(ctx, "SCIENTIFIC CENTER", curr_y, '50px "Tinos"', COLOR_NAVY) + 30;

                    curr_y = drawCenteredText(ctx, "CERTIFICATE", curr_y, '260px "Cinzel"', COLOR_GOLD) + 10;
                    
                    let gradLine = ctx.createLinearGradient(CANVAS_WIDTH/2 - 400, 0, CANVAS_WIDTH/2 + 400, 0);
                    gradLine.addColorStop(0, "rgba(184, 134, 11, 0)");
                    gradLine.addColorStop(0.5, COLOR_GOLD);
                    gradLine.addColorStop(1, "rgba(184, 134, 11, 0)");
                    ctx.fillStyle = gradLine;
                    ctx.fillRect(CANVAS_WIDTH/2 - 400, curr_y, 800, 4);
                    curr_y += 40; 

                    curr_y = drawCenteredText(ctx, "OF APPRECIATION", curr_y, '100px "Tinos"', COLOR_LIGHT_GOLD) + 90;
                    
                    curr_y = drawCenteredText(ctx, "Is Proudly Awarded To", curr_y, '70px "Tinos"', COLOR_NAVY) + 40;

                    let nameFont = getFittedFont(ctx, studentName, CANVAS_WIDTH - 800, 280, 100, "Great Vibes"); 
                    curr_y = drawCenteredText(ctx, studentName, curr_y, nameFont, COLOR_GOLD) + 20;

                    ctx.strokeStyle = COLOR_LIGHT_GOLD;
                    ctx.lineWidth = 4;
                    ctx.beginPath();
                    ctx.moveTo(CANVAS_WIDTH / 2 - 600, curr_y);
                    ctx.lineTo(CANVAS_WIDTH / 2 + 600, curr_y);
                    ctx.stroke();
                    curr_y += 50; 

                    let bodyText = "In Recognition Of Outstanding Commitment, Exceptional Dedication, And High Mastery Demonstrated Throughout The Training Program. This Certificate Stands As A Testament To Your Hard Work, Perseverance, And Dedication To Continuous Learning.";
                    // Decreased font from 85px to 70px
                    curr_y = drawWrappedText(ctx, bodyText, curr_y, '70px "Tinos"', COLOR_NAVY, CANVAS_WIDTH - 700, CANVAS_WIDTH) + 50;

                    // Title Case Course Name
                    let displayCourseName = courseName.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');
                    let courseFont = getFittedFont(ctx, displayCourseName, CANVAS_WIDTH - 600, 140, 80, "Tinos");
                    curr_y = drawCenteredText(ctx, displayCourseName, curr_y, courseFont, COLOR_GOLD) + 80;
                    
                    // 4. Grid Layout for Bottom Signatures (5 Columns)
                    const gridWidth = CANVAS_WIDTH - (mInner * 2); 
                    const colWidth = gridWidth / 5; 
                    const startX = mInner; 
                    
                    const col1_center = startX + colWidth * 0.5; // CEO
                    const col3_center = startX + colWidth * 2.5; // Date
                    const col5_center = startX + colWidth * 4.5; // Instructor

                    let sig_y = CANVAS_HEIGHT - mInner - 130; // Move down heavily
                    let sig_line_width = 500; // Increased line width

                    // Bottom: CEO (Far Left)
                    if (ceoLogoImg) {
                        let s = scaleImage(ceoLogoImg, 650, 350); // Bigger stamp
                        let sx = col1_center - s.w/2;
                        ctx.drawImage(ceoLogoImg, sx, sig_y - s.h - 10, s.w, s.h);
                    }
                    
                    ctx.strokeStyle = COLOR_NAVY;
                    ctx.lineWidth = 2;
                    ctx.beginPath(); ctx.moveTo(col1_center - sig_line_width/2, sig_y); ctx.lineTo(col1_center + sig_line_width/2, sig_y); ctx.stroke();
                    
                    let cur_l_y = sig_y + 15;
                    cur_l_y = drawCenteredUnderLine(ctx, CEO_NAME, col1_center - sig_line_width/2, sig_line_width, cur_l_y, '95px "Great Vibes"', COLOR_GOLD) + 10;
                    drawCenteredUnderLine(ctx, "CEO", col1_center - sig_line_width/2, sig_line_width, cur_l_y, '50px "Tinos"', COLOR_NAVY); // Bigger description

                    // Date (Center) - Move down heavily but keep size
                    drawCenteredText(ctx, `Date: ${issueMonth}, ${issueYear}`, sig_y + 150, '65px "Tinos"', COLOR_NAVY, col3_center * 2);

                    // Instructor (Far Right)
                    if (useInstructor) {
                        let inst_x = col5_center;
                        
                        if (instLogoImg) {
                            let s = scaleImage(instLogoImg, 650, 350); // Bigger stamp
                            let sx = inst_x - s.w/2;
                            ctx.drawImage(instLogoImg, sx, sig_y - s.h - 10, s.w, s.h);
                        }
                        
                        ctx.beginPath(); ctx.moveTo(inst_x - sig_line_width/2, sig_y); ctx.lineTo(inst_x + sig_line_width/2, sig_y); ctx.stroke();
                        
                        let cur_i_y = sig_y + 15;
                        cur_i_y = drawCenteredUnderLine(ctx, instName, inst_x - sig_line_width/2, sig_line_width, cur_i_y, '95px "Great Vibes"', COLOR_GOLD) + 10;
                        drawCenteredUnderLine(ctx, instTitle, inst_x - sig_line_width/2, sig_line_width, cur_i_y, '50px "Tinos"', COLOR_NAVY); // Bigger description
                    }
                }


                // Add to PDF
                const imgData = canvas.toDataURL("image/jpeg", 0.98); // Ultra high quality
                if (i > 0) doc.addPage([CANVAS_WIDTH, CANVAS_HEIGHT], 'landscape');
                doc.addImage(imgData, 'JPEG', 0, 0, CANVAS_WIDTH, CANVAS_HEIGHT);
            }
            
            progressStatus.textContent = "جاري تجميع الملف (PDF)...";
            await delay(500);

            const pdfBlob = doc.output('blob');
            const url = URL.createObjectURL(pdfBlob);

            progressContainer.classList.add('d-none');
            resultContainer.classList.remove('d-none');
            headerSection.classList.remove('d-none');
            
            document.getElementById('res-total-certs').textContent = totalStudents;
            document.getElementById('res-total-students').textContent = totalStudents;
            document.getElementById('res-time-taken').textContent = `${((Date.now() - startTime) / 1000).toFixed(1)}s`;
            
            const downloadBtn = document.getElementById('download-btn');
            downloadBtn.href = url;
            downloadBtn.download = `Certificates_${courseName.replace(/\s+/g, '_')}.pdf`;
            
        } catch (err) {
            console.error(err);
            showToast(`حدث خطأ: ${err.message}`, 'danger');
            progressContainer.classList.add('d-none');
            mainDashboard.classList.remove('d-none');
            headerSection.classList.remove('d-none');
            submitBtn.disabled = false;
        }
    });

    document.getElementById('restart-btn').addEventListener('click', () => {
        document.getElementById('result-container').classList.add('d-none');
        document.getElementById('main-dashboard').classList.remove('d-none');
        document.getElementById('process-btn').disabled = false;
        document.getElementById('progress-bar').style.width = '0%';
        document.getElementById('progress-bar').textContent = '0%';
    });
});
