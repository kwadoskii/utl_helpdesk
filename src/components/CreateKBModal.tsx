"use client";

import { useState, SyntheticEvent } from "react";
import { X, Plus, Trash2 } from "lucide-react";
import { apiRequest } from "@/api";
import { AccountInfo } from "@azure/msal-browser";
import { useMsal } from "@azure/msal-react";

interface CreateKBModalProps {
  open: boolean;
  onClose: () => void;
  onCreated: () => void;
}

export default function CreateKBModal({ open, onClose, onCreated }: CreateKBModalProps) {
  const { instance: msalInstance, accounts } = useMsal();
  const instance = msalInstance as import("@azure/msal-browser").PublicClientApplication;

  const account = (accounts[0] || instance.getActiveAccount()) as AccountInfo | null;

  const [title, setTitle] = useState("");
  const [category, setCategory] = useState("General");
  const [summary, setSummary] = useState("");
  const [body, setBody] = useState("");
  const [tags, setTags] = useState<string[]>([]);
  const [tagInput, setTagInput] = useState("");

  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  if (!open) return null;

  const addTag = () => {
    const value = tagInput.trim();

    if (value && !tags.includes(value)) {
      setTags((prev) => [...prev, value]);
    }

    setTagInput("");
  };

  const removeTag = (tag: string) => {
    setTags((prev) => prev.filter((item) => item !== tag));
  };

  const createKnowledgeBase = async (e: SyntheticEvent) => {
    e.preventDefault();

    if (!account) {
      setError("Your Microsoft account could not be identified.");
      return;
    }

    if (!title.trim() || !summary.trim() || !body.trim()) {
      setError("Please complete all required fields.");
      return;
    }

    setSaving(true);
    setError("");

    try {
      await apiRequest(instance, account, "/knowledge", {
        method: "POST",
        body: JSON.stringify({
          title: title.trim(),
          category,
          summary: summary.trim(),
          body: body.trim(),
          published: true,
          tags,
        }),
      });

      setTitle("");
      setCategory("General");
      setSummary("");
      setBody("");
      setTags([]);
      setTagInput("");

      onCreated();
      onClose();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not create knowledge base article.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="modal-backdrop">
      <form className="modal" onSubmit={createKnowledgeBase} onClick={(e) => e.stopPropagation()}>
        <div className="modal-head">
          <div>
            <h2>Create Knowledge Base Article</h2>
            <p>Add a new article to the MIS knowledge base.</p>
          </div>
          <button type="button" className="icon-btn" onClick={onClose}>
            <X size={18} />
          </button>
        </div>

        <label>
          Article title *
          <input
            value={title}
            onChange={(e) => setTitle(e.target.value)}
            placeholder="e.g. How to connect to the company VPN"
            required
          />
        </label>

        <label>
          Category
          <select value={category} onChange={(e) => setCategory(e.target.value)}>
            {[
              "General",
              "Printer",
              "Network",
              "Email",
              "Access",
              "Hardware",
              "Software",
              "Security",
              "Other",
            ].map((x) => (
              <option key={x}>{x}</option>
            ))}
          </select>
        </label>

        <label>
          Summary *
          <textarea
            rows={3}
            value={summary}
            onChange={(e) => setSummary(e.target.value)}
            placeholder="Briefly describe what this article explains..."
            required
          />
        </label>

        <label>
          Article content *
          <textarea
            rows={7}
            value={body}
            onChange={(e) => setBody(e.target.value)}
            placeholder="Write the full article or step-by-step guide..."
            required
          />
        </label>

        <div className="">
          <label>Tags</label>

          <div>
            <div className="flex gap-1">
              <input
                value={tagInput}
                onChange={(e) => setTagInput(e.target.value)}
                placeholder="Enter a tag"
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    e.preventDefault();
                    addTag();
                  }
                }}
                className="outline-brand !p-3 !py-1 !text-sm"
              />
              <button
                type="button"
                className="bg-[--surface] text-[--text] border-[--border] border rounded-lg text-xs cursor-pointer py-1.5 px-3"
                onClick={addTag}
              >
                <Plus size={16} /> Add
              </button>
            </div>
          </div>
        </div>

        <div className="kb-tags">
          {tags.map((tag) => (
            <span
              key={tag}
              className="!bg-brand/90 text-white mt-2 p-1 px-2 items-center rounded-lg hover:!bg-brand"
            >
              {tag}
              <button
                type="button"
                className="ml-2"
                onClick={() => removeTag(tag)}
                aria-label={`Remove ${tag}`}
              >
                <Trash2 size={12} />
              </button>
            </span>
          ))}
        </div>

        <div className="modal-actions">
          <button type="button" className="secondary-btn" onClick={onClose}>
            Cancel
          </button>
          <button
            className="primary-btn"
            disabled={saving || !title.trim() || !summary.trim() || !body.trim()}
          >
            {saving ? (
              "Saving..."
            ) : (
              <>
                <Plus size={16} /> Save
              </>
            )}
          </button>
        </div>
      </form>
    </div>
  );
}
