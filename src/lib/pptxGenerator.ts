import pptxgen from 'pptxgenjs';

export async function exportMarkdownToPptx(markdown: string, topic: string, subject: string, grade: string): Promise<Buffer> {
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_169';
  
  // Set theme colors
  const tealColor = '1F4E79';
  const darkTeal = '16365C';
  const lightTeal = 'F2F6FA';
  const accentColor = 'E07A5F';
  const textColor = '2C3E50';
  const white = 'FFFFFF';
  
  // Parse sections of markdown
  const lines = markdown.split(/\r?\n/);
  let currentSection = '';
  const sections: { [key: string]: string[] } = {
    title: [],
    objectives: [],
    equipment: [],
    act1: [],
    act2: [],
    act3: [],
    act4: []
  };
  
  // Basic parsing
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    
    if (trimmed.startsWith('# ')) {
      sections.title.push(trimmed.substring(2));
    } else if (trimmed.includes('Mục tiêu') || trimmed.startsWith('## I.')) {
      currentSection = 'objectives';
    } else if (trimmed.includes('Thiết bị') || trimmed.startsWith('## II.')) {
      currentSection = 'equipment';
    } else if (trimmed.includes('Tiến trình') || trimmed.startsWith('## III.')) {
      currentSection = 'process';
    } else if (trimmed.includes('Xác định vấn đề') || trimmed.includes('Khởi động') || trimmed.includes('Hoạt động 1:')) {
      currentSection = 'act1';
    } else if (trimmed.includes('Hình thành kiến thức') || trimmed.includes('Khám phá') || trimmed.includes('Hoạt động 2:')) {
      currentSection = 'act2';
    } else if (trimmed.includes('Luyện tập') || trimmed.includes('Hoạt động 3:')) {
      currentSection = 'act3';
    } else if (trimmed.includes('Vận dụng') || trimmed.includes('Hoạt động 4:')) {
      currentSection = 'act4';
    } else if (trimmed.startsWith('## ') || trimmed.startsWith('### ')) {
      // General transition, do not capture headings as bullets
    } else if (currentSection && currentSection !== 'process') {
      // Remove markdown bold markings and list markers for slide bullets
      const cleanLine = trimmed
        .replace(/\*\*/g, '')
        .replace(/^\s*[-*+]\s*/, '')
        .replace(/^\s*\d+\.\s*/, '');
      if (cleanLine.length > 3) {
        sections[currentSection].push(cleanLine);
      }
    }
  }

  // Helper to add standard header/background to content slides
  const addContentSlide = (titleText: string) => {
    const slide = pptx.addSlide();
    // Top banner
    slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: '100%', h: 1.1, fill: { color: tealColor } });
    slide.addText(titleText, { x: 0.5, y: 0.2, w: 12, h: 0.7, fontSize: 24, bold: true, color: white, fontFace: 'Arial' });
    // Background color
    slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 1.1, w: '100%', h: 6.4, fill: { color: lightTeal } });
    // Footer
    slide.addText(`Môn: ${subject} - Lớp: ${grade} | App soạn GA NLS HUY 2`, { x: 0.5, y: 7.1, w: 12.3, h: 0.3, fontSize: 10, italic: true, color: '7F8C8D', fontFace: 'Arial' });
    return slide;
  };

  // 1. Title Slide
  const slide1 = pptx.addSlide();
  // Background
  slide1.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: '100%', h: '100%', fill: { color: lightTeal } });
  // Left accent bar
  slide1.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 0.4, h: '100%', fill: { color: accentColor } });
  // Main title container box
  slide1.addShape(pptx.shapes.RECTANGLE, { x: 1.2, y: 1.5, w: 10.9, h: 4.4, fill: { color: white }, line: { color: tealColor, width: 2 } });
  
  slide1.addText(subject.toUpperCase(), { x: 1.8, y: 2.0, w: 9.8, h: 0.5, fontSize: 18, color: accentColor, bold: true, fontFace: 'Arial' });
  slide1.addText(topic || sections.title[0] || 'BÀI GIẢNG ĐIỆN TỬ', { x: 1.8, y: 2.6, w: 9.8, h: 2.0, fontSize: 32, color: darkTeal, bold: true, fontFace: 'Arial', verticalAlign: 'middle' });
  slide1.addText(`Lớp: ${grade} | Giáo án Chuẩn Khung Năng Lực Số`, { x: 1.8, y: 4.9, w: 9.8, h: 0.5, fontSize: 14, color: textColor, italic: true, fontFace: 'Arial' });

  // 2. Slide 2: Objectives
  const slide2 = addContentSlide('MỤC TIÊU BÀI HỌC');
  let objText = sections.objectives.slice(0, 7).map(item => `• ${item}`).join('\n\n');
  if (!objText) objText = '• Kiến thức: Nắm vững các nội dung trọng tâm bài dạy.\n\n• Năng lực: Phát triển năng lực số, tự chủ, tự học.\n\n• Phẩm chất: Rèn luyện tính tự giác, tinh thần trách nhiệm.';
  slide2.addText(objText, { x: 0.8, y: 1.6, w: 11.7, h: 5.0, fontSize: 15, color: textColor, fontFace: 'Arial', lineSpacing: 22 });

  // 3. Slide 3: Equipment & Digital Resources
  const slide3 = addContentSlide('THIẾT BỊ DẠY HỌC & HỌC LIỆU SỐ');
  let eqText = sections.equipment.slice(0, 7).map(item => `• ${item}`).join('\n\n');
  if (!eqText) eqText = '• Giáo viên: Máy tính, máy chiếu, bài giảng điện tử, tài liệu số.\n\n• Học sinh: Sách giáo khoa, vở ghi, thiết bị thông minh (nếu có).';
  slide3.addText(eqText, { x: 0.8, y: 1.6, w: 11.7, h: 5.0, fontSize: 15, color: textColor, fontFace: 'Arial', lineSpacing: 22 });

  // 4. Slide 4: Lesson Activities Overview
  const slide4 = addContentSlide('TIẾN TRÌNH DẠY HỌC');
  const cardW = 2.7;
  const cardH = 4.0;
  const startX = 0.7;
  const gap = 0.3;
  
  const activities = [
    { title: 'HĐ 1: Khởi động', desc: 'Xác định nhiệm vụ, vấn đề cần giải quyết thông qua câu hỏi thực tiễn.' },
    { title: 'HĐ 2: Khám phá', desc: 'Hình thành kiến thức, kĩ năng mới qua các hoạt động và học liệu số.' },
    { title: 'HĐ 3: Luyện tập', desc: 'Hệ thống hóa kiến thức và làm các bài tập củng cố, thực hành.' },
    { title: 'HĐ 4: Vận dụng', desc: 'Giải quyết các vấn đề thực tế liên quan và mở rộng nghiên cứu.' }
  ];
  
  activities.forEach((act, idx) => {
    const xPos = startX + idx * (cardW + gap);
    // Card background
    slide4.addShape(pptx.shapes.RECTANGLE, { x: xPos, y: 1.8, w: cardW, h: cardH, fill: { color: white }, line: { color: tealColor, width: 1 } });
    // Card header
    slide4.addShape(pptx.shapes.RECTANGLE, { x: xPos, y: 1.8, w: cardW, h: 0.8, fill: { color: tealColor } });
    slide4.addText(act.title, { x: xPos, y: 1.9, w: cardW, h: 0.6, fontSize: 14, bold: true, color: white, align: 'center', fontFace: 'Arial' });
    // Card body
    slide4.addText(act.desc, { x: xPos + 0.1, y: 2.8, w: cardW - 0.2, h: cardH - 1.2, fontSize: 12, color: textColor, align: 'center', fontFace: 'Arial' });
  });

  // 5. Slides for each activity details
  const activitySlides = [
    { key: 'act1', title: 'HOẠT ĐỘNG 1: KHỞI ĐỘNG (XÁC ĐỊNH VẤN ĐỀ)' },
    { key: 'act2', title: 'HOẠT ĐỘNG 2: HÌNH THÀNH KIẾN THỨC MỚI' },
    { key: 'act3', title: 'HOẠT ĐỘNG 3: LUYỆN TẬP - CỦNG CỐ' },
    { key: 'act4', title: 'HOẠT ĐỘNG 4: VẬN DỤNG - MỞ RỘNG' }
  ];
  
  activitySlides.forEach(actInfo => {
    const slide = addContentSlide(actInfo.title);
    let details = sections[actInfo.key].slice(0, 6).map(item => `• ${item}`).join('\n\n');
    if (!details) {
      details = '• Giáo viên giao nhiệm vụ học tập thông qua công cụ số.\n\n• Học sinh thảo luận nhóm, phân tích dữ liệu.\n\n• Kết quả được trình bày và chốt kiến thức.';
    }
    slide.addText(details, { x: 0.8, y: 1.6, w: 11.7, h: 5.0, fontSize: 15, color: textColor, fontFace: 'Arial', lineSpacing: 22 });
  });

  // Generate buffer
  const buffer = await pptx.write('nodebuffer');
  return buffer as Buffer;
}
