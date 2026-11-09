export default function StatusPill({ value }: { value: string }) {
  const c = value.toLowerCase().replaceAll(" ", "-");
  return <span className={`pill pill-${c}`}>{value}</span>;
}
