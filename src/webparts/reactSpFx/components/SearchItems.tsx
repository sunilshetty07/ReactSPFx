import * as React from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DetailsList, IColumn, PrimaryButton, TextField } from "@fluentui/react";

interface ISearchProps {
  context: WebPartContext;
  listName: string;
  pageSize?: number; // optional override
}

const SearchItems: React.FC<ISearchProps> = ({ context, listName, pageSize = 100 }) => {
  const [query, setQuery] = React.useState<string>("");
  const [results, setResults] = React.useState<any[]>([]);
  const [loading, setLoading] = React.useState<boolean>(false);
  const [nextLink, setNextLink] = React.useState<string | null>(null);
  const [noMore, setNoMore] = React.useState<boolean>(false);

  const webUrl = context.pageContext.web.absoluteUrl;

  // Columns for the Fluent UI DetailsList
  const columns: IColumn[] = [
    { key: "id", name: "ID", fieldName: "Id", minWidth: 50, maxWidth: 70 },
    {
      key: "title",
      name: "Title",
      fieldName: "Title",
      minWidth: 200,
      onRender: (item: any) => {
        // Opens the default SharePoint display form in a new tab
        const listDispUrl = `${webUrl}/Lists/${encodeURIComponent(listName)}/DispForm.aspx?ID=${item.Id}`;
        return (
          <a href={listDispUrl} target="_blank" rel="noopener noreferrer">
            {item.Title}
          </a>
        );
      }
    }
  ];

  // Helper: build initial REST URL for substring search on Title
  const buildInitialUrl = (searchText: string) => {
    // escape single quotes in query for OData
    const safe = (searchText || "").replace(/'/g, "''");
    // OData substringof: substringof('text', Title)
    const filter = `substringof('${safe}', Title)`;
    // encode the filter portion
    const encodedFilter = encodeURIComponent(filter);
    // Note: listName is used raw inside getbytitle('...') — this is common practice
    return `${webUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title&$filter=${encodedFilter}&$top=${pageSize}`;
  };

  // Generic fetch page function using SPHttpClient.
  // Accepts a REST URL (initial or nextLink).
  const fetchPage = async (url: string) => {
    // request minimal metadata to receive @odata.nextLink (works in modern SPO)
    const headers = {
      "Accept": "application/json;odata=minimalmetadata",
        "odata-version": ""
    };

    const resp: SPHttpClientResponse = await context.spHttpClient.get(url, SPHttpClient.configurations.v1, { headers });
    const json = await resp.json();

    // support multiple response shapes:
    // - modern: json.value and json['@odata.nextLink']
    // - older: json.d.results and json.d.__next
    let items: any[] = [];
    let next: string | null = null;

    if (json.value && Array.isArray(json.value)) {
      items = json.value;
      next = (json as any)["@odata.nextLink"] || null;
    } else if (json.d && json.d.results) {
      items = json.d.results;
      next = (json.d as any).__next || null;
    } else {
      // fallback: if the response is an array already
      if (Array.isArray(json)) {
        items = json;
      }
    }

    return { items, next };
  };

  // Main search function: reset=true for initial search; false to load more
  const performSearch = async (reset = true) => {
    if (!query && reset) {
      // nothing to search
      return;
    }

    setLoading(true);

    try {
      if (reset) {
        setResults([]);
        setNextLink(null);
        setNoMore(false);
      }

      const url = reset ? buildInitialUrl(query) : (nextLink as string);
      if (!url) {
        setLoading(false);
        return;
      }

      const { items, next } = await fetchPage(url);

      setResults(prev => (reset ? items : [...prev, ...items]));
      if (next) {
        setNextLink(next);
        setNoMore(false);
      } else {
        setNextLink(null);
        setNoMore(true);
      }
    } catch (e) {
      console.error("Search failed:", e);
      // graceful fallback
      if (reset) setResults([]);
      setNextLink(null);
      setNoMore(true);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ padding: 16 }}>
      <h3>Search list: {listName}</h3>

      <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 12 }}>
        <TextField
          placeholder="Enter text to search in Title..."
          value={query}
          onChange={(_, v) => setQuery(v || "")}
          styles={{ root: { minWidth: 320 } }}
        />
        <PrimaryButton text="Search" onClick={() => performSearch(true)} disabled={!query || loading} />
      </div>

      {/* Searching state */}
      {loading && results.length === 0 && <div>Searching...</div>}

      {/* No results */}
      {!loading && results.length === 0 && query && <div>No items found</div>}

      {/* Results table */}
      {results.length > 0 && (
        <div>
          <DetailsList items={results} columns={columns} setKey="search" layoutMode={0} />
          {/* Load more */}
          {!noMore && (
            <div style={{ marginTop: 12 }}>
              <PrimaryButton text={loading ? "Loading..." : "Load more"} onClick={() => performSearch(false)} disabled={loading} />
            </div>
          )}
        </div>
      )}
    </div>
  );
};

export default SearchItems;
