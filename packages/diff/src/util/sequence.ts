/** Pairwise alignment of two sequences using longest common subsequence.
 *  Returns the matched index pairs and the unmatched left/right indices.
 *
 *  Used to align paragraphs (DOCX) and slides (PPTX) when there is no stable
 *  ID, so we can detect inserts/removes vs. modifications. */

export interface SequenceAlignment {
  /** `[leftIndex, rightIndex]` pairs in ascending order. */
  matches: Array<[number, number]>;
  /** Indices in left not present in any match. Sorted ascending. */
  removed: number[];
  /** Indices in right not present in any match. Sorted ascending. */
  added: number[];
}

/** LCS alignment driven by an equality predicate. O(n*m) time/space — fine for
 *  the typical document sizes we deal with (≤ a few thousand paragraphs). */
export function alignSequences<L, R>(
  left: readonly L[],
  right: readonly R[],
  equals: (a: L, b: R) => boolean,
): SequenceAlignment {
  const n = left.length;
  const m = right.length;
  const dp: number[][] = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));

  for (let i = 1; i <= n; i++) {
    for (let j = 1; j <= m; j++) {
      if (equals(left[i - 1], right[j - 1])) {
        dp[i][j] = dp[i - 1][j - 1] + 1;
      } else {
        dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
      }
    }
  }

  const matches: Array<[number, number]> = [];
  const matchedLeft = new Set<number>();
  const matchedRight = new Set<number>();

  let i = n;
  let j = m;
  while (i > 0 && j > 0) {
    if (equals(left[i - 1], right[j - 1])) {
      matches.push([i - 1, j - 1]);
      matchedLeft.add(i - 1);
      matchedRight.add(j - 1);
      i--;
      j--;
    } else if (dp[i - 1][j] >= dp[i][j - 1]) {
      i--;
    } else {
      j--;
    }
  }
  matches.reverse();

  const removed: number[] = [];
  for (let k = 0; k < n; k++) if (!matchedLeft.has(k)) removed.push(k);
  const added: number[] = [];
  for (let k = 0; k < m; k++) if (!matchedRight.has(k)) added.push(k);

  return { matches, removed, added };
}
