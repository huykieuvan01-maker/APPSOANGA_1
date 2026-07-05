import React, { useState, useEffect } from 'react';
import { 
  BookOpen, Sparkles, Loader2, FileUp, Info, FileText, 
  Presentation, Download, Edit3, Eye, Save, Trash2, 
  Gamepad2, FileQuestion, Settings 
} from 'lucide-react';
import ReactMarkdown from 'react-markdown';
import remarkMath from 'remark-math';
import rehypeKatex from 'rehype-katex';
import { LessonPlanRequest, ApiResponse } from './types';
import { cn } from './lib/utils';

export default function App() {
  const [activeTab, setActiveTab] = useState<'create' | 'refactor'>('create');
  
  const [topic, setTopic] = useState('');
  const [grade, setGrade] = useState('');
  const [subject, setSubject] = useState('');
  const [oldContent, setOldContent] = useState('');

  const [apiKey, setApiKey] = useState(() => localStorage.getItem('gemini_api_key') || '');
  const [aiModel, setAiModel] = useState(() => localStorage.getItem('gemini_ai_model') || 'gemini-3-flash-preview');
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);

  useEffect(() => {
    const savedKey = localStorage.getItem('gemini_api_key');
    if (!savedKey) {
      setIsSettingsOpen(true);
    }
  }, []);

  const [tempApiKey, setTempApiKey] = useState(apiKey);
  const [tempAiModel, setTempAiModel] = useState(aiModel);

  useEffect(() => {
    if (isSettingsOpen) {
      setTempApiKey(apiKey);
      setTempAiModel(aiModel);
    }
  }, [isSettingsOpen, apiKey, aiModel]);

  const handleSaveSettings = () => {
    localStorage.setItem('gemini_api_key', tempApiKey.trim());
    localStorage.setItem('gemini_ai_model', tempAiModel);
    setApiKey(tempApiKey.trim());
    setAiModel(tempAiModel);
    setIsSettingsOpen(false);
  };
  
  const [isLoading, setIsLoading] = useState(false);
  const [result, setResult] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  const [isReadingFile, setIsReadingFile] = useState(false);
  const [fileError, setFileError] = useState<string | null>(null);

  const [isEditingMode, setIsEditingMode] = useState(false);
  const [isExportingDocx, setIsExportingDocx] = useState(false);
  const [isExportingPptx, setIsExportingPptx] = useState(false);
  const [isGeneratingExtras, setIsGeneratingExtras] = useState(false);

  interface SavedPlan {
    id: string;
    topic: string;
    subject: string;
    grade: string;
    content: string;
    timestamp: number;
  }

  const [history, setHistory] = useState<SavedPlan[]>(() => {
    try {
      const saved = localStorage.getItem('lesson_plans_history');
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  const saveToHistory = () => {
    if (!result) return;
    const newPlan: SavedPlan = {
      id: Date.now().toString(),
      topic: topic || 'Giáo án không tên',
      subject: subject || 'Môn học',
      grade: grade || 'Lớp',
      content: result,
      timestamp: Date.now()
    };
    const updated = [newPlan, ...history.filter(p => p.topic !== newPlan.topic || p.subject !== newPlan.subject)];
    setHistory(updated);
    localStorage.setItem('lesson_plans_history', JSON.stringify(updated));
    alert('Đã lưu giáo án vào kho cá nhân thành công!');
  };

  const loadFromHistory = (plan: SavedPlan) => {
    setTopic(plan.topic);
    setSubject(plan.subject);
    setGrade(plan.grade);
    setResult(plan.content);
    setError(null);
  };

  const deleteFromHistory = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    const updated = history.filter(item => item.id !== id);
    setHistory(updated);
    localStorage.setItem('lesson_plans_history', JSON.stringify(updated));
  };

  const handleExportDocx = async () => {
    if (!result) return;
    setIsExportingDocx(true);
    try {
      const response = await fetch('/api/export-docx', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ markdown: result })
      });
      if (!response.ok) throw new Error('Không thể xuất tệp Word.');
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Giao_an_${(topic || 'chua_dat_ten').replace(/\s+/g, '_')}.docx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    } catch (err: any) {
      alert(err.message || 'Lỗi khi tải file Word.');
    } finally {
      setIsExportingDocx(false);
    }
  };

  const handleExportPptx = async () => {
    if (!result) return;
    setIsExportingPptx(true);
    try {
      const response = await fetch('/api/export-pptx', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ 
          markdown: result,
          topic: topic,
          subject: subject,
          grade: grade
        })
      });
      if (!response.ok) throw new Error('Không thể xuất tệp PowerPoint.');
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Bai_giang_${(topic || 'chua_dat_ten').replace(/\s+/g, '_')}.pptx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    } catch (err: any) {
      alert(err.message || 'Lỗi khi tải file PowerPoint.');
    } finally {
      setIsExportingPptx(false);
    }
  };

  const fetchWithFallback = async (promptText: string): Promise<string> => {
    const MODELS_LIST = ['gemini-3-flash-preview', 'gemini-3-pro-preview', 'gemini-2.5-flash'];
    const modelsToTry = [
      aiModel,
      ...MODELS_LIST.filter(m => m !== aiModel)
    ];

    let lastError = '';
    for (const currentModel of modelsToTry) {
      try {
        const response = await fetch('/api/generate', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({ 
            prompt: promptText, 
            apiKey: apiKey, 
            model: currentModel 
          }),
        });

        const data: ApiResponse = await response.json();

        if (!response.ok) {
          throw new Error(data.error || `Lỗi API từ model ${currentModel}`);
        }

        if (data.result) {
          return data.result;
        }
      } catch (err: any) {
        console.warn(`Model ${currentModel} thất bại: ${err.message}. Đang thử model tiếp theo...`);
        lastError = err.message;
      }
    }
    throw new Error(lastError || 'Tất cả các model đều thất bại. Đã dừng do lỗi.');
  };

  const generateSupplemental = async (type: 'exam' | 'game') => {
    if (!result) return;
    setIsGeneratingExtras(true);
    setError(null);
    
    let promptText = '';
    if (type === 'exam') {
      promptText = `Dựa trên giáo án môn ${subject}, lớp ${grade}, chủ đề "${topic}" sau đây, hãy thiết kế một đề kiểm tra 15 phút (gồm 10 câu trắc nghiệm khách quan có đáp án rõ ràng và 1 câu tự luận kèm thang điểm). Trình bày bằng Markdown:\n\n${result}`;
    } else {
      promptText = `Dựa trên giáo án môn ${subject}, lớp ${grade}, chủ đề "${topic}" sau đây, hãy viết mã nguồn một mini-game ôn tập tương tác (ví dụ game trắc nghiệm đố vui có tính điểm và hiệu ứng). Game phải được thiết kế dưới dạng MỘT TỆP HTML DUY NHẤT (Single-page HTML) chứa đầy đủ mã CSS và JavaScript để giáo viên có thể mở trực tiếp trên trình duyệt hoặc trình chiếu. Hãy trả về toàn bộ mã nguồn HTML trong khối code \`\`\`html \`\`\` và hướng dẫn cách mở game:\n\n${result}`;
    }

    try {
      const generatedResult = await fetchWithFallback(promptText);
      setResult(generatedResult);
      setIsEditingMode(false);
    } catch (err: any) {
      setError(err.message);
    } finally {
      setIsGeneratingExtras(false);
    }
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsReadingFile(true);
    setFileError(null);

    const formData = new FormData();
    formData.append('file', file);

    try {
      const response = await fetch('/api/parse-file', {
        method: 'POST',
        body: formData,
      });

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.error || 'Lỗi khi trích xuất nội dung từ tệp.');
      }

      if (data.text) {
        setOldContent(data.text);
      }
    } catch (err: any) {
      setFileError(err.message);
    } finally {
      setIsReadingFile(false);
      e.target.value = '';
    }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setIsLoading(true);
    setError(null);
    setResult(null);

    let prompt = '';
    if (activeTab === 'create') {
      prompt = `Hãy tạo một giáo án chuẩn khung năng lực số cho môn ${subject}, lớp ${grade}, chủ đề: "${topic}". Vui lòng tuân thủ nghiêm ngặt các quy định về khung năng lực số và cấu trúc Markdown, LaTeX như trong System Instruction.`;
    } else {
      prompt = `Hãy nâng cấp và số hóa giáo án cũ sau đây thuộc môn ${subject}, lớp ${grade}, chủ đề: "${topic}". Vui lòng tuân thủ nghiêm ngặt các quy định về khung năng lực số và cấu trúc Markdown, LaTeX như trong System Instruction. Nội dung cũ:\n\n${oldContent}`;
    }

    try {
      const generatedResult = await fetchWithFallback(prompt);
      setResult(generatedResult);
    } catch (err: any) {
      setError(err.message);
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans selection:bg-blue-100 selection:text-blue-900">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10">
        <div className="max-w-6xl mx-auto px-4 sm:px-6 h-16 flex items-center justify-between">
          <div className="flex items-center gap-2 text-blue-600">
            <BookOpen className="w-6 h-6" />
            <span className="text-lg font-bold tracking-tight text-slate-900">App soạn GA NLS HUY 2</span>
          </div>
          <div className="flex items-center gap-3">
            <button 
              type="button"
              onClick={() => setIsSettingsOpen(true)}
              className={cn(
                "flex items-center gap-2 px-3 py-1.5 rounded-lg text-xs font-semibold transition-all cursor-pointer border shadow-sm",
                !apiKey 
                  ? "bg-red-50 text-red-600 border-red-200 animate-pulse hover:bg-red-100 font-bold" 
                  : "bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100"
              )}
            >
              <Settings className="w-3.5 h-3.5 animate-spin-slow" />
              {!apiKey ? "Lấy API key để sử dụng app" : "Cài đặt API & Model"}
            </button>
            <div className="text-xs font-semibold text-slate-500 bg-slate-100 px-3 py-1.5 rounded-full hidden sm:block">
              Chuẩn Khung Năng Lực Số
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-6xl mx-auto px-4 sm:px-6 py-8 flex flex-col lg:flex-row gap-8">
        
        {/* Left Column: Input Form */}
        <div className="w-full lg:w-1/3 flex flex-col gap-6">
          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
            <div className="flex border-b border-slate-200">
              <button
                className={cn(
                  "flex-1 py-3 text-sm font-medium transition-colors",
                  activeTab === 'create' ? "text-blue-600 border-b-2 border-blue-600 bg-blue-50/50" : "text-slate-500 hover:text-slate-700 hover:bg-slate-50"
                )}
                onClick={() => setActiveTab('create')}
              >
                Tạo Mới Giáo Án
              </button>
              <button
                className={cn(
                  "flex-1 py-3 text-sm font-medium transition-colors",
                  activeTab === 'refactor' ? "text-blue-600 border-b-2 border-blue-600 bg-blue-50/50" : "text-slate-500 hover:text-slate-700 hover:bg-slate-50"
                )}
                onClick={() => setActiveTab('refactor')}
              >
                Nâng Cấp / Số Hóa
              </button>
            </div>
            
            <div className="p-5">
              <form onSubmit={handleSubmit} className="flex flex-col gap-4">
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1">Môn học</label>
                  <input
                    type="text"
                    required
                    value={subject}
                    onChange={(e) => setSubject(e.target.value)}
                    placeholder="VD: Toán, Vật lý, Ngữ văn..."
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1">Lớp</label>
                  <input
                    type="text"
                    required
                    value={grade}
                    onChange={(e) => setGrade(e.target.value)}
                    placeholder="VD: 10, 11, 12..."
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1">Tên bài học / Chủ đề</label>
                  <input
                    type="text"
                    required
                    value={topic}
                    onChange={(e) => setTopic(e.target.value)}
                    placeholder="VD: Phương trình bậc 2..."
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all"
                  />
                </div>

                {activeTab === 'refactor' && (
                  <div className="flex flex-col gap-4">
                    <div>
                      <label className="block text-sm font-medium text-slate-700 mb-1.5 flex items-center gap-1.5">
                        <FileUp className="w-4 h-4 text-slate-400" />
                        Tải lên tệp giáo án cũ
                      </label>
                      <div className="border-2 border-dashed border-slate-200 rounded-xl p-4 text-center bg-slate-50 hover:bg-blue-50/20 hover:border-blue-300 transition-all cursor-pointer relative group">
                        <input
                          type="file"
                          accept=".docx,.pdf"
                          onChange={handleFileUpload}
                          className="absolute inset-0 opacity-0 cursor-pointer w-full h-full"
                          disabled={isReadingFile}
                        />
                        <div className="flex flex-col items-center gap-1.5">
                          <FileUp className="w-8 h-8 text-slate-400 group-hover:text-blue-500 transition-colors" />
                          <span className="text-xs font-semibold text-slate-600 group-hover:text-blue-600 transition-colors">
                            Kéo thả hoặc click để tải lên tệp
                          </span>
                          <span className="text-[10px] text-slate-400">
                            Hỗ trợ tệp .docx (Word) hoặc .pdf
                          </span>
                        </div>
                      </div>
                      
                      {isReadingFile && (
                        <div className="flex items-center justify-center gap-2 text-xs text-blue-600 font-medium py-2.5 mt-1">
                          <Loader2 className="w-4 h-4 animate-spin" />
                          <span>Đang trích xuất nội dung từ tệp...</span>
                        </div>
                      )}

                      {fileError && (
                        <div className="text-xs text-red-600 bg-red-50 p-2.5 rounded-lg border border-red-100 font-medium mt-2">
                          {fileError}
                        </div>
                      )}
                    </div>

                    <div>
                      <label className="block text-sm font-medium text-slate-700 mb-1">
                        Nội dung giáo án cũ (Đã trích xuất hoặc tự nhập)
                      </label>
                      <textarea
                        required
                        value={oldContent}
                        onChange={(e) => setOldContent(e.target.value)}
                        placeholder="Nội dung tệp sẽ được trích xuất tại đây, hoặc bạn có thể tự dán/nhập nội dung..."
                        className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm min-h-[180px] resize-y focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all font-mono"
                      />
                    </div>
                  </div>
                )}

                <button
                  type="submit"
                  disabled={isLoading}
                  className="mt-2 w-full bg-blue-600 hover:bg-blue-700 text-white font-medium py-2.5 px-4 rounded-lg flex items-center justify-center gap-2 transition-all disabled:opacity-70 disabled:cursor-not-allowed shadow-sm shadow-blue-500/20"
                >
                  {isLoading ? (
                    <Loader2 className="w-5 h-5 animate-spin" />
                  ) : (
                    <Sparkles className="w-5 h-5" />
                  )}
                  {isLoading ? 'Đang xử lý...' : 'Tạo Giáo Án Số'}
                </button>
              </form>
            </div>
          </div>
          
          
          <div className="bg-blue-50 rounded-xl p-4 border border-blue-100 flex gap-3 text-sm text-blue-800">
            <Info className="w-5 h-5 shrink-0 text-blue-600" />
            <p>Hệ thống tự động tích hợp ít nhất 2 chỉ số năng lực số và xuất công thức Toán/Lý/Hóa dưới định dạng LaTeX, sẵn sàng copy sang Word (MathType).</p>
          </div>

          {/* Kho Giáo Án Cá Nhân */}
          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden flex flex-col">
            <div className="px-5 py-3.5 border-b border-slate-100 bg-slate-50/50 flex items-center justify-between">
              <h3 className="font-semibold text-slate-800 flex items-center gap-1.5 text-xs">
                <BookOpen className="w-4 h-4 text-blue-500" />
                Kho giáo án cá nhân ({history.length})
              </h3>
            </div>
            <div className="p-3 max-h-[300px] overflow-y-auto flex flex-col gap-2">
              {history.length === 0 ? (
                <p className="text-xs text-slate-400 text-center py-4">Chưa có giáo án nào được lưu. Bấm "Lưu kho" ở kết quả để lưu trữ.</p>
              ) : (
                history.map((plan) => (
                  <div 
                    key={plan.id}
                    onClick={() => loadFromHistory(plan)}
                    className="group flex items-center justify-between p-2 rounded-lg border border-slate-100 hover:border-blue-200 hover:bg-blue-50/20 transition-all cursor-pointer text-left"
                  >
                    <div className="flex flex-col gap-0.5 overflow-hidden pr-2">
                      <span className="text-xs font-semibold text-slate-700 truncate">{plan.topic}</span>
                      <span className="text-[10px] text-slate-400 truncate">{plan.subject} - Lớp {plan.grade}</span>
                    </div>
                    <button 
                      onClick={(e) => deleteFromHistory(plan.id, e)}
                      className="text-slate-400 hover:text-red-500 p-1 rounded hover:bg-slate-100 transition-colors opacity-0 group-hover:opacity-100 shrink-0"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>
                ))
              )}
            </div>
          </div>
        </div>

        {/* Right Column: Result Viewer */}
        <div className="w-full lg:w-2/3">
          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 h-full min-h-[600px] flex flex-col overflow-hidden">
            <div className="px-5 py-4 border-b border-slate-100 bg-slate-50/50 flex flex-wrap items-center justify-between gap-3">
              <h2 className="font-semibold text-slate-800 flex items-center gap-1.5 text-sm">
                <FileText className="w-4 h-4 text-blue-500" />
                Kết quả Trình xuất
              </h2>
              {result && (
                <div className="flex flex-wrap items-center gap-2">
                  {/* Toggle Preview/Edit */}
                  <button
                    type="button"
                    onClick={() => setIsEditingMode(!isEditingMode)}
                    className={cn(
                      "text-xs font-medium px-2.5 py-1.5 rounded-md border flex items-center gap-1.5 shadow-sm transition-colors cursor-pointer",
                      isEditingMode ? "bg-blue-50 border-blue-200 text-blue-600" : "bg-white border-slate-200 text-slate-600 hover:bg-slate-50"
                    )}
                  >
                    {isEditingMode ? (
                      <>
                        <Eye className="w-3.5 h-3.5" />
                        <span>Xem trước</span>
                      </>
                    ) : (
                      <>
                        <Edit3 className="w-3.5 h-3.5" />
                        <span>Sửa trực tiếp</span>
                      </>
                    )}
                  </button>

                  {/* Save to library */}
                  <button
                    type="button"
                    onClick={saveToHistory}
                    className="text-xs font-medium bg-white border border-slate-200 px-2.5 py-1.5 rounded-md shadow-sm text-slate-600 hover:text-blue-600 hover:border-blue-200 flex items-center gap-1.5 transition-colors cursor-pointer"
                  >
                    <Save className="w-3.5 h-3.5" />
                    <span>Lưu kho</span>
                  </button>

                  {/* Export Word */}
                  <button
                    type="button"
                    onClick={handleExportDocx}
                    disabled={isExportingDocx}
                    className="text-xs font-medium bg-white border border-slate-200 px-2.5 py-1.5 rounded-md shadow-sm text-slate-600 hover:text-blue-600 hover:border-blue-200 flex items-center gap-1.5 transition-colors disabled:opacity-50 cursor-pointer"
                  >
                    {isExportingDocx ? (
                      <Loader2 className="w-3.5 h-3.5 animate-spin" />
                    ) : (
                      <FileText className="w-3.5 h-3.5 text-blue-500" />
                    )}
                    <span>Xuất Word (.docx)</span>
                  </button>

                  {/* Export PowerPoint */}
                  <button
                    type="button"
                    onClick={handleExportPptx}
                    disabled={isExportingPptx}
                    className="text-xs font-medium bg-white border border-slate-200 px-2.5 py-1.5 rounded-md shadow-sm text-slate-600 hover:text-blue-600 hover:border-blue-200 flex items-center gap-1.5 transition-colors disabled:opacity-50 cursor-pointer"
                  >
                    {isExportingPptx ? (
                      <Loader2 className="w-3.5 h-3.5 animate-spin" />
                    ) : (
                      <Presentation className="w-3.5 h-3.5 text-orange-500" />
                    )}
                    <span>Tạo Slide (.pptx)</span>
                  </button>

                  {/* Copy Markdown */}
                  <button 
                    type="button"
                    onClick={() => {
                      navigator.clipboard.writeText(result);
                      alert('Đã sao chép nội dung vào Clipboard!');
                    }}
                    className="text-xs font-medium bg-blue-600 text-white px-3 py-1.5 rounded-md shadow-sm hover:bg-blue-700 flex items-center gap-1.5 transition-colors cursor-pointer"
                  >
                    <span>Sao chép</span>
                  </button>
                </div>
              )}
            </div>
            
            <div className="p-6 md:p-8 flex-1 overflow-auto">
              {isLoading ? (
                <div className="h-full flex flex-col items-center justify-center text-slate-400 gap-4">
                  <div className="w-16 h-16 relative">
                    <div className="absolute inset-0 border-4 border-slate-100 rounded-full"></div>
                    <div className="absolute inset-0 border-4 border-blue-500 rounded-full border-t-transparent animate-spin"></div>
                  </div>
                  <p className="font-medium animate-pulse">AI đang thiết kế bài giảng...</p>
                </div>
              ) : error ? (
                <div className="bg-red-50 text-red-700 p-4 rounded-xl border border-red-100 font-medium">
                  {error}
                </div>
              ) : result ? (
                isEditingMode ? (
                  <div className="h-full flex flex-col min-h-[500px]">
                    <textarea
                      value={result}
                      onChange={(e) => setResult(e.target.value)}
                      className="w-full flex-1 p-4 font-mono text-xs border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 resize-y min-h-[500px] bg-slate-50 text-slate-800 leading-relaxed"
                    />
                  </div>
                ) : (
                  <div className="flex flex-col gap-6">
                    {/* Supplemental Activities Quick Actions */}
                    <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 flex flex-col sm:flex-row items-center justify-between gap-4">
                      <div className="flex items-center gap-2">
                        <Sparkles className="w-5 h-5 text-blue-500" />
                        <div>
                          <h4 className="text-xs font-bold text-slate-800">Hoạt động bổ trợ giảng dạy</h4>
                          <p className="text-[10px] text-slate-500">Tự động thiết kế bộ công cụ học tập từ giáo án này</p>
                        </div>
                      </div>
                      <div className="flex items-center gap-2 w-full sm:w-auto">
                        <button
                          type="button"
                          onClick={() => generateSupplemental('exam')}
                          disabled={isGeneratingExtras}
                          className="flex-1 sm:flex-initial text-xs font-semibold bg-white border border-slate-200 hover:border-blue-200 hover:text-blue-600 px-3 py-2 rounded-lg flex items-center justify-center gap-1.5 transition-colors disabled:opacity-50 cursor-pointer"
                        >
                          <FileQuestion className="w-4 h-4 text-blue-500" />
                          Tạo Đề Kiểm Tra (15p)
                        </button>
                        <button
                          type="button"
                          onClick={() => generateSupplemental('game')}
                          disabled={isGeneratingExtras}
                          className="flex-1 sm:flex-initial text-xs font-semibold bg-white border border-slate-200 hover:border-blue-200 hover:text-blue-600 px-3 py-2 rounded-lg flex items-center justify-center gap-1.5 transition-colors disabled:opacity-50 cursor-pointer"
                        >
                          <Gamepad2 className="w-4 h-4 text-emerald-500" />
                          Tạo Trò Chơi Ôn Tập (HTML)
                        </button>
                      </div>
                    </div>

                    {isGeneratingExtras && (
                      <div className="flex items-center justify-center gap-2 py-4 text-sm text-blue-600 font-medium animate-pulse">
                        <Loader2 className="w-5 h-5 animate-spin" />
                        <span>AI đang thiết kế hoạt động bổ trợ bài giảng...</span>
                      </div>
                    )}

                    <div className="prose prose-slate prose-blue max-w-none prose-headings:font-semibold prose-a:text-blue-600 prose-p:leading-relaxed prose-pre:bg-slate-800 prose-pre:text-slate-100">
                      <ReactMarkdown
                        remarkPlugins={[remarkMath]}
                        rehypePlugins={[rehypeKatex]}
                      >
                        {result}
                      </ReactMarkdown>
                    </div>
                  </div>
                )
              ) : (
                <div className="h-full flex flex-col items-center justify-center text-slate-400 gap-4 opacity-60">
                  <BookOpen className="w-16 h-16" />
                  <p className="font-medium text-center">Hoàn thiện thông tin bên trái để <br/> bắt đầu khởi tạo giáo án.</p>
                </div>
              )}
            </div>
          </div>
        </div>

      </main>

      {/* Settings Modal */}
      {isSettingsOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-lg overflow-hidden border border-slate-100 flex flex-col max-h-[90vh] animate-in zoom-in-95 duration-200">
            {/* Header */}
            <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between">
              <div className="flex items-center gap-2 text-slate-800 font-semibold">
                <Settings className="w-5 h-5 text-blue-600" />
                <h2 className="text-sm font-bold">Cài đặt API & Model</h2>
              </div>
              {apiKey && (
                <button 
                  type="button"
                  onClick={() => setIsSettingsOpen(false)} 
                  className="p-1.5 text-slate-400 hover:text-slate-600 hover:bg-slate-50 rounded-lg transition-colors cursor-pointer"
                >
                  <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                  </svg>
                </button>
              )}
            </div>
            
            {/* Content */}
            <div className="p-6 overflow-y-auto space-y-5 text-left">
              {/* Models Cards list */}
              <div className="space-y-2">
                <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider">1. Chọn Model AI</label>
                <div className="grid grid-cols-1 gap-2.5">
                  {[
                    { id: 'gemini-3-flash-preview', name: 'Gemini 3 Flash', desc: 'Mới nhất, siêu nhanh và tiết kiệm' },
                    { id: 'gemini-3-pro-preview', name: 'Gemini 3 Pro', desc: 'Thông minh nhất, xử lý các giáo án phức tạp' },
                    { id: 'gemini-2.5-flash', name: 'Gemini 2.5 Flash', desc: 'Phiên bản ổn định, tối ưu chi phí' }
                  ].map((m) => (
                    <div 
                      key={m.id}
                      onClick={() => setTempAiModel(m.id)}
                      className={cn(
                        "p-3 rounded-xl border-2 cursor-pointer transition-all flex flex-col gap-0.5 text-left",
                        tempAiModel === m.id 
                          ? "border-blue-600 bg-blue-50/40" 
                          : "border-slate-200 hover:border-slate-300 hover:bg-slate-50/50"
                      )}
                    >
                      <div className="flex items-center justify-between">
                        <span className="font-semibold text-slate-900 text-xs">{m.name}</span>
                        <span className="text-[9px] text-slate-500 font-mono bg-slate-100 px-1.5 py-0.5 rounded">{m.id}</span>
                      </div>
                      <p className="text-[11px] text-slate-500">{m.desc}</p>
                    </div>
                  ))}
                </div>
              </div>

              {/* API Key Input */}
              <div className="space-y-2">
                <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider">2. Nhập Gemini API Key</label>
                <div className="relative">
                  <input
                    type="password"
                    value={tempApiKey}
                    onChange={(e) => setTempApiKey(e.target.value)}
                    placeholder="Nhập API Key của bạn (AIzaSy...)"
                    className="block w-full px-3 py-2 border border-slate-200 rounded-lg text-xs outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all font-mono"
                  />
                </div>
                <div className="bg-blue-50/50 text-blue-800 p-3 rounded-xl border border-blue-100/50 text-xs leading-relaxed">
                  Bạn chưa có API Key? Quá đơn giản! Hãy{' '}
                  <a 
                    href="https://aistudio.google.com/api-keys" 
                    target="_blank" 
                    rel="noopener noreferrer" 
                    className="text-blue-600 font-semibold hover:underline inline-flex items-center gap-1"
                  >
                    Lấy API key hoàn toàn miễn phí tại đây (Google AI Studio)
                  </a>
                </div>
              </div>
            </div>

            {/* Footer */}
            <div className="px-6 py-4 bg-slate-50 border-t border-slate-100 flex justify-end gap-2">
              {apiKey && (
                <button 
                  type="button"
                  onClick={() => setIsSettingsOpen(false)}
                  className="px-4 py-2 text-xs font-semibold text-slate-600 bg-white border border-slate-200 rounded-lg hover:bg-slate-50 shadow-sm cursor-pointer"
                >
                  Hủy
                </button>
              )}
              <button 
                type="button"
                onClick={handleSaveSettings}
                disabled={!tempApiKey.trim()}
                className="px-5 py-2 text-xs font-semibold text-white bg-blue-600 rounded-lg shadow-sm hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed transition-colors cursor-pointer"
              >
                Lưu cài đặt
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
