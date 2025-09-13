import * as React from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DetailsList, IColumn, PrimaryButton, TextField } from "@fluentui/react";

interface ISearchProps {
  context: WebPartContext;
  listName: string;
}

const SearchItems: React.FC<ISearchProps> = ({ context, listName }) => {
  const [query, setQuery] = React.useState<string>("");
  const [results, setResults] = React.useState<any[]>([]);
  const [loading, setLoading] = React.useState<boolean>(false);

  const webUrl = context.pageContext.web.absoluteUrl;

  const columns: IColumn[] = [
    { key: "id", name: "ID", fieldName: "Id", minWidth: 50, maxWidth: 70 },
    {
      key: "title",
      name: "Title",
      fieldName: "Title",
      minWidth: 200,
      onRender: (item: any) => {
        const listDispUrl = `${webUrl}/Lists/${encodeURIComponent(listName)}/DispForm.aspx?ID=${item.Id}`;
        return (
          <a href={listDispUrl} target="_blank" rel="noopener noreferrer">
            {item.Title}
          </a>
        );
      }
    }
  ];

  // Fetch one batch of 5000 items and apply local substring search
  const fetchBatch = async (url: string, accumulatedResults: any[] = []): Promise<any[]> => {
    const headers = { Accept: "application/json;odata=minimalmetadata", "odata-version": "" };
    const resp: SPHttpClientResponse = await context.spHttpClient.get(url, SPHttpClient.configurations.v1, { headers });
    const json = await resp.json();

    let batchItems: any[] = [];
    let nextLink: string | null = null;

    if (json.value && Array.isArray(json.value)) {
      batchItems = json.value.filter((i: { Title: string; }) => i.Title?.toLowerCase().includes(query.toLowerCase()));
      nextLink = json["odata.nextLink"] || null;
      console.log("NextLink from @odata.nextLink:", nextLink);
    } else if (json.d?.results) {
      batchItems = json.d.results.filter((i: { Title: string; }) => i.Title?.toLowerCase().includes(query.toLowerCase()));
      nextLink = json.d.__next || null;
      console.log(json.json.d?.results);
      console.log("NextLink from d.__next:", nextLink);
    }

    // Append the current batch’s matches to accumulated results
    const updatedResults = [...accumulatedResults, ...batchItems];
    setResults(updatedResults); // update UI progressively

    // If there’s a next batch, fetch recursively
    if (nextLink) {
      return fetchBatch(nextLink, updatedResults);
    } else {
      return updatedResults;
    }
  };

  const performSearch = async () => {
    if (!query) return;
    setLoading(true);
    setResults([]);

    try {
      const initialUrl = `${webUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id,Title&$top=5000`;
      await fetchBatch(initialUrl);
    } catch (err) {
      console.error("Search failed", err);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ padding: 16 }}>
      <h3>Search list: {listName}</h3>
      <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 12 }}>
        <TextField
          placeholder="Enter text to search..."
          value={query}
          onChange={(_, v) => setQuery(v || "")}
          styles={{ root: { minWidth: 320 } }}
        />
        <PrimaryButton text="Search" onClick={performSearch} disabled={!query || loading} />
      </div>

      {loading && <div>Searching...</div>}
      {!loading && results.length === 0 && query && <div>No items found</div>}

      {results.length > 0 && (
        <DetailsList
          items={results}
          columns={columns}
          setKey="search"
          onShouldVirtualize={() => false} // render all items
        />
      )}
    </div>
  );
};

export default SearchItems;
