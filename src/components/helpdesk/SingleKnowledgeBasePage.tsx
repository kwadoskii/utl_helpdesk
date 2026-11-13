"use client";

import { useEffect, useState } from "react";
import Link from "next/link";
import { useParams } from "next/navigation";
// import ReactMarkdown from "react-markdown";
// import remarkGfm from "remark-gfm";
import {
  ArrowLeft,
  BookOpen,
  CalendarDays,
  Clock3,
  Tag,
  AlertCircle,
  RefreshCw,
} from "lucide-react";
import { apiRequest } from "@/api";
import { Article } from "@/lib/types";
import { AccountInfo } from "@azure/msal-browser";
import { useMsal } from "@azure/msal-react";
import { relativeDate } from "@/lib/relativeDate";
import { toast } from "sonner";

export default function SingleKnowledgeBasePage() {
  const params = useParams<{ slug: string }>();
  const slug = params.slug;

  const { instance: msalInstance, accounts } = useMsal();
  const instance = msalInstance as import("@azure/msal-browser").PublicClientApplication;
  const account = (accounts[0] || instance.getActiveAccount()) as AccountInfo;

  const [article, setArticle] = useState<Article | null>(null);
  const [loading, setLoading] = useState(true);
  // const [error, setError] = useState("");

  const fetchArticle = async () => {
    try {
      setLoading(true);
      // setError("");

      // Uses the same API helper as the Knowledge Base listing page.
      const data = await apiRequest<Article>(
        instance,
        account,
        `/knowledge/${encodeURIComponent(slug)}`,
      );

      if (!data || !data.published) {
        setArticle(null);
        // setError("This article could not be found.");
        toast.error("This article could not be found.");
        return;
      }

      setArticle(data);
    } catch (err) {
      console.error("Failed to load knowledge article:", err);
      // setError("Unable to load this article. Please try again.");
      toast.error("Unable to load this article. Please try again.");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (slug) fetchArticle();
  }, [slug]);

  if (loading) {
    return (
      <div className="page">
        <div className="kb-article-state">
          <RefreshCw className="kb-spin" size={25} />
          <p>Loading article ...</p>
        </div>
      </div>
    );
  }

  if (!article) {
    return (
      <div className="page">
        {/* <Link href="/knowledge-base" className="kb-back-link">
          <ArrowLeft size={15} />
          Back to Knowledge Base
        </Link> */}

        <div className="kb-article-state">
          <div className="kb-state-icon">
            <AlertCircle size={25} />
          </div>
          <h2>Article unavailable</h2>
          <p>The requested article was not found.</p>
          <Link href="/knowledge-base" className="primary-btn">
            Browse Knowledge Base
          </Link>
        </div>
      </div>
    );
  }

  return (
    <div className="page kb-article-page">
      {/* Back navigation */}
      <Link href="/knowledge-base" className="kb-back-link">
        <ArrowLeft size={15} />
        Back to Knowledge Base
      </Link>

      <div className="kb-article-layout">
        <main className="kb-article-main">
          {/* Article heading */}
          <div className="kb-article-header">
            <div className="kb-article-category">
              <BookOpen size={13} />
              {article.category || "General"}
            </div>

            <h1>{article.title}</h1>

            {article.summary && <p className="kb-article-summary">{article.summary}</p>}

            <div className="kb-article-meta">
              <span>
                <CalendarDays size={14} />
                Published {article.createdAt ? relativeDate(article.createdAt) : "-"}
              </span>

              <span>
                <Clock3 size={14} />
                Knowledge Base Article
              </span>
            </div>
          </div>

          {/* Article content */}
          <article className="kb-article-content">
            {/* <ReactMarkdown remarkPlugins={[remarkGfm]}> */}
            {article.body || ""}
            {/* </ReactMarkdown> */}
          </article>

          {/* Article tags */}
          {article.tags && article.tags.length > 0 && (
            <div className="kb-article-tags">
              <div className="kb-tags-heading">
                <Tag size={14} />
                Article tags
              </div>

              <div className="kb-tags-list">
                {article.tags.map((tag) => (
                  <span className="kb-article-tag" key={tag}>
                    {tag}
                  </span>
                ))}
              </div>
            </div>
          )}

          {/* Footer */}
          <div className="kb-article-footer">
            <div>
              <BookOpen size={17} />
              <span>End of article</span>
            </div>

            <Link href="/knowledge-base" className="text-btn">
              Explore more articles <span>→</span>
            </Link>
          </div>
        </main>

        {/* Right-hand information panel */}
        <aside className="kb-article-sidebar">
          <div className="kb-article-side-card">
            <div className="kb-side-icon">
              <BookOpen size={19} />
            </div>

            <h3>Knowledge Base</h3>
            <p>
              Find helpful guides, solutions and information to help you resolve common IT issues.
            </p>

            <Link href="/knowledge-base" className="kb-side-link">
              Browse all articles <span>→</span>
            </Link>
          </div>

          <div className="kb-article-side-card kb-help-card">
            <h3>Still need help?</h3>
            <p>If this article did not resolve your issue, contact the MIS team.</p>

            <Link href="/my-tickets" className="primary-btn kb-help-btn">
              View My Tickets
            </Link>
          </div>
        </aside>
      </div>
    </div>
  );
}
