import express from 'express';
import path from 'path';
import { createServer as createViteServer } from 'vite';
import { GoogleGenAI } from '@google/genai';
import multer from 'multer';
import mammoth from 'mammoth';
import pdfParse from 'pdf-parse';
import { exportMarkdownToDocx } from './src/lib/docxExporter';
import { exportMarkdownToPptx } from './src/lib/pptxGenerator';

const upload = multer({ storage: multer.memoryStorage() });

const systemInstruction = `Bạn là một Chuyên gia cao cấp về Phương pháp luận Giáo dục và Chuyển đổi số, có am hiểu sâu sắc về các quy định, thông tư của Bộ Giáo dục và Đào tạo Việt Nam. Vai trò của bạn là trợ lý đắc lực cho giáo viên trong việc thiết kế giáo án (Kế hoạch bài dạy) hiện đại, tuân thủ nghiêm ngặt Khung năng lực số dành cho học sinh và giáo viên. Bạn không chỉ là một công cụ viết lách, mà là một cố vấn chuyên môn có khả năng phân tích dữ liệu cũ và chuyển đổi chúng thành các sản phẩm giáo dục chuẩn hóa thế kỷ 21.

1. Tạo mới giáo án chuẩn hóa: Hỗ trợ giáo viên xây dựng giáo án từ ý tưởng sơ khai, đảm bảo đầy đủ các thành phần: Mục tiêu bài học (Kiến thức, Năng lực, Phẩm chất), Thiết bị dạy học, và Tiến trình dạy học theo các công văn hướng dẫn mới nhất (ví dụ: Công văn 5512/BGDĐT).
2. Số hóa và Nâng cấp (Refactor): Phân tích nội dung giáo án cũ để trích xuất cốt lõi chuyên môn, sau đó tự động tái cấu trúc và bổ sung các chỉ số năng lực số phù hợp với chương trình giáo dục phổ thông mới.
3. Chuẩn hóa kỹ thuật MathType: Đảm bảo toàn bộ công thức Toán học, Vật lý, Hóa học được soạn thảo bằng định dạng trung gian (LaTeX/MathML) có khả năng chuyển đổi hoàn hảo sang đối tượng MathType có thể chỉnh sửa được trên Microsoft Word.
4. Tối ưu hóa năng lực số: Tích hợp các hoạt động ứng dụng CNTT vào giáo án một cách tự nhiên, giúp học sinh phát triển năng lực khai thác, quản lý và sáng tạo trên môi trường số.

Guidelines & Rules:
- Tuân thủ khung năng lực: Mọi mục tiêu bài học phải được tham chiếu dựa trên Khung năng lực số của Bộ GD&ĐT. Các hoạt động học tập phải thể hiện rõ việc ứng dụng công cụ số.
- Quy trình xử lý tệp cũ: Phân tích, nhận diện, đề xuất chỉnh sửa: Giữ nguyên giá trị chuyên môn, thay đổi phương pháp sư phạm.
- Quy định về Công thức (Quan trọng):
  - TẤT CẢ công thức toán học/khoa học phải được viết bằng cú pháp LaTeX (ví dụ: \`$x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}$\`).
  - Không được sử dụng định dạng văn bản thuần túy hoặc hình ảnh cho công thức.
  - Đảm bảo cấu trúc LaTeX chuẩn để hệ thống backend có thể biên dịch chính xác sang đối tượng MathType khi xuất file Word.
- Cấu trúc giáo án: Phải bao gồm các phần: Tên bài học, Mục tiêu, Thiết bị dạy học và học liệu số, Tiến trình dạy học (Gồm 4 hoạt động: Xác định vấn đề; Hình thành kiến thức; Luyện tập; Vận dụng).
- Kiểm chứng: Bạn phải tự kiểm tra xem giáo án đã tích hợp đủ ít nhất 2 chỉ số năng lực số hay chưa trước khi phản hồi.

Tone & Persona: Chuyên nghiệp, học thuật, chính xác nhưng vẫn đảm bảo tính gợi mở và sáng tạo. Tiếng Việt chuẩn mực sư phạm.
Output Format: Trình bày theo cấu trúc Markdown rõ ràng.

1. Phần tổng quan: Tóm tắt thay đổi hoặc các điểm nhấn năng lực số.
2. Nội dung chi tiết:
  - I. Mục tiêu
  - II. Thiết bị & Học liệu số
  - III. Tiến trình dạy học
3. Khu vực công thức: Các công thức lồng ghép trong nội dung phải ở dạng LaTeX chuẩn.
4. Ghi chú kỹ thuật: Xác nhận về việc sẵn sàng xuất file Word với MathType editable.`;

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json());

  // API endpoint to extract text from docx/pdf
  app.post('/api/parse-file', upload.single('file'), async (req: any, res: any) => {
    try {
      if (!req.file) {
        return res.status(400).json({ error: 'Không tìm thấy tệp được tải lên.' });
      }

      const fileBuffer = req.file.buffer;
      const originalName = req.file.originalname;
      const fileExtension = originalName.split('.').pop()?.toLowerCase();

      let text = '';

      if (fileExtension === 'docx') {
        const result = await mammoth.extractRawText({ buffer: fileBuffer });
        text = result.value;
      } else if (fileExtension === 'doc') {
        return res.status(400).json({ error: 'Định dạng .doc cũ không được hỗ trợ. Vui lòng chuyển đổi sang .docx hoặc .pdf rồi tải lên.' });
      } else if (fileExtension === 'pdf') {
        const data = await pdfParse(fileBuffer);
        text = data.text;
      } else {
        return res.status(400).json({ error: 'Định dạng tệp không được hỗ trợ. Chỉ chấp nhận tệp .docx hoặc .pdf.' });
      }

      res.json({ text });
    } catch (error: any) {
      console.error('Lỗi khi trích xuất file:', error);
      res.status(500).json({ error: error.message || 'Có lỗi xảy ra khi xử lý tệp.' });
    }
  });

  app.post('/api/generate', async (req, res) => {
    try {
      const { prompt, apiKey, model } = req.body;
      if (!prompt) {
        return res.status(400).json({ error: 'Prompt is required' });
      }

      const finalApiKey = apiKey || process.env.GEMINI_API_KEY;
      if (!finalApiKey) {
        return res.status(400).json({ error: 'GEMINI_API_KEY is not configured. Please provide it in the request or configure on the server.' });
      }

      const ai = new GoogleGenAI({ 
        apiKey: finalApiKey,
        httpOptions: { headers: { 'User-Agent': 'aistudio-build' } }
      });

      const finalModel = model || 'gemini-3.1-pro-preview';

      const response = await ai.models.generateContent({
        model: finalModel,
        contents: prompt,
        config: {
          systemInstruction: systemInstruction,
          temperature: 0.7,
        }
      });

      res.json({ result: response.text });
    } catch (error: any) {
      console.error('Error generating content:', error);
      res.status(500).json({ error: error.message || 'An error occurred during generation.' });
    }
  });

  app.post('/api/export-docx', async (req, res) => {
    try {
      const { markdown } = req.body;
      if (!markdown) {
        return res.status(400).json({ error: 'Nội dung giáo án trống.' });
      }

      const buffer = await exportMarkdownToDocx(markdown);
      
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      res.setHeader('Content-Disposition', 'attachment; filename=giaoan.docx');
      res.send(buffer);
    } catch (error: any) {
      console.error('Lỗi khi xuất Word:', error);
      res.status(500).json({ error: error.message || 'Lỗi khi sinh file Word.' });
    }
  });

  app.post('/api/export-pptx', async (req, res) => {
    try {
      const { markdown, topic, subject, grade } = req.body;
      if (!markdown) {
        return res.status(400).json({ error: 'Nội dung giáo án trống.' });
      }

      const buffer = await exportMarkdownToPptx(markdown, topic || '', subject || '', grade || '');
      
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
      res.setHeader('Content-Disposition', 'attachment; filename=baigiang.pptx');
      res.send(buffer);
    } catch (error: any) {
      console.error('Lỗi khi xuất PowerPoint:', error);
      res.status(500).json({ error: error.message || 'Lỗi khi sinh file PowerPoint.' });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
  });
}

startServer();
