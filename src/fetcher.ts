export function buildSharePointDownloadUrl(sharepointUrl: string): string {
  const url = new URL(sharepointUrl);
  url.search = "";
  url.searchParams.set("download", "1");
  return url.toString();
}

function extractCookies(headers: Headers): string[] {
  const rawSetCookies =
    typeof headers.getSetCookie === "function"
      ? headers.getSetCookie()
      : headers.get("set-cookie")
        ? [headers.get("set-cookie")!]
        : [];

  const cookies: string[] = [];
  for (const header of rawSetCookies) {
    const firstPart = header.split(";")[0]?.trim();
    if (firstPart) {
      cookies.push(firstPart);
    }
  }
  return cookies;
}

export async function downloadTimetableXlsx(sharepointUrl: string): Promise<ArrayBuffer> {
  let currentUrl = buildSharePointDownloadUrl(sharepointUrl);
  const cookieJar = new Map<string, string>();
  const maxRedirects = 5;

  for (let hop = 0; hop <= maxRedirects; hop++) {
    const cookieHeader = Array.from(cookieJar.entries())
      .map(([k, v]) => `${k}=${v}`)
      .join("; ");

    const response = await fetch(currentUrl, {
      method: "GET",
      redirect: "manual",
      headers: {
        "User-Agent": "Mozilla/5.0 (compatible; NSBM-Timetable-Sync/1.0)",
        ...(cookieHeader ? { Cookie: cookieHeader } : {}),
      },
    });

    for (const cookieStr of extractCookies(response.headers)) {
      const eqIdx = cookieStr.indexOf("=");
      if (eqIdx > 0) {
        const name = cookieStr.slice(0, eqIdx).trim();
        const val = cookieStr.slice(eqIdx + 1).trim();
        cookieJar.set(name, val);
      }
    }

    if (response.status >= 300 && response.status < 400) {
      const location = response.headers.get("location");
      if (!location) {
        throw new Error(`Redirect (${response.status}) without Location header from ${currentUrl}`);
      }
      currentUrl = new URL(location, currentUrl).toString();
      continue;
    }

    if (!response.ok) {
      throw new Error(`Failed to download SharePoint file: HTTP ${response.status} ${response.statusText}`);
    }

    const buffer = await response.arrayBuffer();
    const bytes = new Uint8Array(buffer);
    if (
      bytes.length < 4 ||
      bytes[0] !== 0x50 ||
      bytes[1] !== 0x4b ||
      bytes[2] !== 0x03 ||
      bytes[3] !== 0x04
    ) {
      const preview = new TextDecoder().decode(bytes.slice(0, 200));
      throw new Error(
        `Downloaded content from SharePoint is not a valid .xlsx file (first bytes: ${JSON.stringify(preview)})`
      );
    }

    return buffer;
  }

  throw new Error(`Exceeded maximum redirects (${maxRedirects}) downloading SharePoint timetable`);
}
