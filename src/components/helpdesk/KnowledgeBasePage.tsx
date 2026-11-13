"use client";

import { apiRequest } from "@/api";
import { AccountInfo } from "@azure/msal-browser";
import { useMsal } from "@azure/msal-react";
import { BookOpen, Search } from "lucide-react";
import { useEffect, useState } from "react";
import NewButton from "../NewTicketButton";
import CreateKBModal from "@/components/CreateKBModal";
import { isSupport } from "@/lib/isSupport";
import { toast } from "sonner";

export default function KnowledgeBasePage() {
  const [articleQuery, setArticleQuery] = useState("");
  const [articles, setArticles] = useState<Article[]>([]);
  const [openedArticle, setOpenedArticle] = useState<Article | null>(null);
  const [loading, setLoading] = useState(false);
  const [showNew, setShowNew] = useState(false);

  const { instance: msalInstance, accounts } = useMsal();
  const instance = msalInstance as import("@azure/msal-browser").PublicClientApplication;
  const account = (accounts[0] || instance.getActiveAccount()) as AccountInfo;

  const loadArticles = async () => {
    if (!account) return;

    setLoading(true);

    try {
      const data = await apiRequest<Article[]>(
        instance,
        account,
        `/knowledge${articleQuery ? `?q=${encodeURIComponent(articleQuery)}` : ""}`,
      );

      setArticles(data);
    } catch (e) {
      // console.error("Could not load knowledge articles:", e);
      toast.error("Could not load knowledge articles.");
    } finally {
      setLoading(false);
    }
  };

  const filtered = articleQuery
    ? articles.filter(
        (a) =>
          a.title?.toLowerCase().includes(articleQuery.toLowerCase().trim()) ||
          a.summary?.toLowerCase().includes(articleQuery.toLowerCase().trim()),
      )
    : articles;

  useEffect(() => {
    void loadArticles();
  }, [account, instance]);

  return (
    <>
      <div className="page">
        <div className="page-heading">
          <div>
            <div className="eyebrow">SELF-SERVICE SUPPORT</div>
            <h1>Knowledge base</h1>
            <p>Search published MIS help article and guides.</p>
          </div>

          <div className="">
            {isSupport(account) && <NewButton text="New" setShowNew={setShowNew} />}
          </div>
        </div>

        <div className="kb-search">
          <Search size={20} />
          <input
            placeholder="Search articles, topics and how-tos..."
            value={articleQuery}
            onChange={(e) => setArticleQuery(e.target.value)}
          />
        </div>
        <div className="kb-grid">
          {filtered.map((a) => (
            <article className="kb-card" key={a._id || a.slug}>
              <div className="kb-icon">
                <BookOpen size={20} />
              </div>
              <span className="kb-category">{a.category}</span>
              <h3>{a.title}</h3>
              <p>{a.summary}</p>
              <a href={`/knowledge-base/${a.slug}`}>
                <button onClick={() => setOpenedArticle(a)}>
                  Read article <span>→</span>
                </button>
              </a>
            </article>
          ))}
        </div>
        {!articles.length && !loading && (
          <div className="empty">
            No published articles found. Added articles would be displayed here.
          </div>
        )}
      </div>

      {showNew && (
        <CreateKBModal open={showNew} onClose={() => setShowNew(false)} onCreated={loadArticles} />
      )}
    </>
  );
}

type Article = {
  _id: string;
  title: string;
  slug: string;
  category: string;
  summary: string;
  body?: string;
  tags?: string[];
};
