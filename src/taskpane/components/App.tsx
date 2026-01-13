import React, { useState } from "react";
import { Box, Button, Checkbox, Heading, Input, Text, VStack } from "@chakra-ui/react";

type SearchResult = {
  text: string;
};

const App = () => {
  const [query, setQuery] = useState("");
  const [caseSensitive, setCaseSensitive] = useState(false);

  const [isLoading, setIsLoading] = useState(false);
  const [results, setResults] = useState<SearchResult[]>([]);
  const [foundCount, setFoundCount] = useState<number | null>(null);
  const [error, setError] = useState<string | null>(null);

  const canSearch = query.trim().length > 0;

  const onSearch = async () => {
    const q = query.trim();
    if (!q) return;

    setIsLoading(true);
    setError(null);

    try {
      await Word.run(async (context: any) => {
        const searchResults = context.document.body.search(q, {
          matchCase: caseSensitive,
        });

        searchResults.load("items/text");
        await context.sync();

        const total = searchResults.items.length;
        setFoundCount(total);

        const top3 = searchResults.items.slice(0, 3).map((r: any) => ({
          text: (r.text || "").trim(),
        }));

        setResults(top3);
        setQuery("");
      });
    } catch (e: any) {
      console.error(e);
      setError(e?.message ?? "Search failed");
      setResults([]);
      setFoundCount(null);
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <Box p={4}>
      <VStack align="stretch" gap={4}>
        <Heading size="md">Word Search</Heading>

        <Input
          placeholder="Search text..."
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && canSearch && !isLoading) onSearch();
          }}
        />

        <Button type="button" onClick={onSearch} disabled={!canSearch || isLoading} colorPalette="blue">
          {isLoading ? "Searching..." : "Search"}
        </Button>

        <Checkbox.Root checked={caseSensitive} onCheckedChange={(details) => setCaseSensitive(!!details.checked)}>
          <Checkbox.HiddenInput />
          <Checkbox.Control />
          <Checkbox.Label>Case sensitive</Checkbox.Label>
        </Checkbox.Root>

        <Box>
          <Text fontSize="sm" color="gray.600" mb={2}>
            Top 3 results:
          </Text>

          {error && (
            <Text fontSize="sm" color="red.500">
              {error}
            </Text>
          )}

          {foundCount !== null && !error && (
            <Text fontSize="sm" color="gray.600" mb={2}>
              Found: {foundCount}
            </Text>
          )}

          {results.length === 0 && !error && (
            <Text fontSize="sm" color="gray.500">
              No results yet...
            </Text>
          )}

          {results.map((r, idx) => (
            <Box key={idx} p={2} borderWidth="1px" borderRadius="md" mb={2}>
              <Text fontSize="sm" fontWeight="600">
                #{idx + 1}
              </Text>
              <Text fontSize="sm">{r.text || "(empty match)"}</Text>
            </Box>
          ))}
        </Box>
      </VStack>
    </Box>
  );
};

export default App;
