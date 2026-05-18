// ─── Instance config ─────────────────────────────────────────────────────────

export type BvInstance = {
  name: string;
  baseUrl: string;
  passkey: string;
  apiversion?: string;
  owner?: string;
};

// ─── Shared envelope ─────────────────────────────────────────────────────────

export type BvError = {
  Message: string;
  Code: string;
};

export type ConversationsResponse<T> = {
  HasErrors: boolean;
  Errors: BvError[];
  TotalResults: number;
  Limit: number;
  Offset: number;
  Results: T[];
  Includes?: Record<string, unknown>;
};

// ─── Products ────────────────────────────────────────────────────────────────

export type RatingDistribution = {
  RatingValue: number;
  Count: number;
};

export type ReviewStatistics = {
  AverageOverallRating: number;
  TotalReviewCount: number;
  RatingsDistribution: RatingDistribution[];
  RecommendedCount?: number;
  NotRecommendedCount?: number;
  OverallRatingRange?: number;
};

export type Product = {
  Id: string;
  Name: string;
  Description?: string;
  CategoryId?: string;
  Active?: boolean;
  Disabled?: boolean;
  BrandExternalId?: string | null;
  Brand?: {
    Id?: string | null;
    Name?: string | null;
  };
  ImageUrl?: string;
  ProductPageUrl?: string;
  UPCs?: string[];
  EANs?: string[];
  FamilyIds?: string[];
  ReviewStatistics?: ReviewStatistics;
  QAStatistics?: {
    TotalQuestionCount: number;
    TotalAnswerCount: number;
  };
};

// ─── Reviews ─────────────────────────────────────────────────────────────────

export type ContextDataValue = {
  Id: string;
  Value: string;
  DimensionLabel?: string;
};

export type Review = {
  Id: string;
  ProductId: string;
  AuthorId: string;
  UserNickname?: string;
  Rating: number;
  Title?: string;
  ReviewText?: string;
  SubmissionTime: string;
  LastModificationTime?: string;
  IsRecommended?: boolean;
  IsSyndicated?: boolean;
  SyndicationSource?: { Name: string; LogoImageUrl?: string };
  Helpfulness?: { Helpfulness: number; TotalPositiveFeedbackCount: number; TotalNegativeFeedbackCount: number };
  ContextDataValues?: Record<string, ContextDataValue>;
  SecondaryRatings?: Record<string, { Id: string; Label: string; Value: number; ValueLabel?: string }>;
  BadgesOrder?: string[];
  TotalCommentCount?: number;
  CID?: string;
};

// ─── Questions & Answers ─────────────────────────────────────────────────────

export type Question = {
  Id: string;
  ProductId: string;
  AuthorId: string;
  UserNickname?: string;
  QuestionSummary: string;
  QuestionDetails?: string;
  SubmissionTime: string;
  TotalAnswerCount: number;
  TotalInappropriateFeedbackCount?: number;
  AnswerIds?: string[];
};

export type Answer = {
  Id: string;
  QuestionId: string;
  ProductId?: string;
  AuthorId: string;
  UserNickname?: string;
  AnswerText: string;
  SubmissionTime: string;
  IsBrandAnswer?: boolean;
  TotalInappropriateFeedbackCount?: number;
};

// ─── Statistics ──────────────────────────────────────────────────────────────

export type ProductStatistics = {
  ProductId: string;
  ReviewStatistics: ReviewStatistics;
  NativeReviewStatistics?: ReviewStatistics;
};

export type StatisticsResult = {
  ProductStatistics: ProductStatistics;
};

