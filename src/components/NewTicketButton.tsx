import { Plus } from "lucide-react";

export default function NewButton({
  setShowNew,
  text = "",
}: {
  setShowNew: (show: boolean) => void;
  text?: string;
}) {
  return (
    <button className="primary-btn" onClick={() => setShowNew(true)}>
      <Plus size={17} /> {text || "New ticket"}
    </button>
  );
}
