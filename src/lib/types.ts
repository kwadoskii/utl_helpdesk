export type TicketItem = {
  _id: string;
  requesterId?: string;
  ticketNumber: string;
  subject: string;
  requesterName?: string;
  requesterEmail?: string;
  department?: string;
  category?: string;
  priority: "Low" | "Medium" | "High" | "Urgent";
  status: Status;
  updatedAt: string;
  assigneeName?: string;
  messages?: Message[];
  createdAt?: string;
  resolvedAt?: string;
};

export type Status = "Open" | "In progress" | "Pending" | "Resolved" | "Closed";

export type Message = {
  _id?: string;
  authorId: string;
  authorName: string;
  authorEmail: string;
  body: string;
  createdAt?: string;
  attachments?: MessageAttachment[];
  isInternalNote?: boolean;
};

export type Article = {
  _id: string;
  title: string;
  slug: string;
  category: string;
  summary: string;
  body?: string;
  content?: string;
  tags?: string[];
  published: boolean;
  createdAt?: string;
};

export type NotificationItem = {
  _id: string;
  recipientId: string;
  type: "ticket_created" | "new_message" | "ticket_closed";
  title: string;
  message: string;
  ticketId:
    | {
        _id: string;
        subject: string;
        status: Status;
      }
    | string;
  isRead: boolean;
  __v?: number;
  createdAt: string;
  updatedAt: string;
};

export type NotificationResponse = {
  notifications: NotificationItem[];
  unreadCount: number;
};

export interface MessageAttachment {
  _id?: string;
  originalName: string;
  storageName: string;
  mimeType: string;
  size: number;
}
