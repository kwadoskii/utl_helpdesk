import { format, formatDistanceToNow } from "date-fns";

export function relativeDate(date: string) {
  const updated = formatDistanceToNow(new Date(date), { addSuffix: true }).replace(
    /about |almost |over |less than /g,
    "",
  );
  // .replace(" seconds ago", "s ago")
  // .replace(" second ago", "s ago")
  // .replace(" minutes ago", "m ago")
  // .replace(" minute ago", "m ago")
  // .replace(" hours ago", "h ago")
  // .replace(" hour ago", "h ago")
  // .replace(" days ago", "d ago")
  // .replace(" day ago", "d ago");

  return updated;
}

export function formatDate(date: string) {
  const formated = format(date, "dd-MMM-yyy  hh:mm a");

  return formated;
}
