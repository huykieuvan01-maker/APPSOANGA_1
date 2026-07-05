export interface LessonPlanRequest {
  topic: string;
  grade: string;
  subject: string;
  oldContent?: string;
}

export interface ApiResponse {
  result?: string;
  error?: string;
}
