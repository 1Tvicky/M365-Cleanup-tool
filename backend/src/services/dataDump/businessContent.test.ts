import { describe, expect, it } from "vitest";
import { createSeededSource } from "./seededRandom.js";
import { buildFolderTree, countFolderNodes, type FolderNode } from "./businessContent.js";

function allNames(tree: FolderNode[]): string[] {
  const names: string[] = [];
  for (const node of tree) {
    names.push(node.name);
    names.push(...allNames(node.children));
  }
  return names;
}

describe("buildFolderTree", () => {
  it("produces professional, non-generic names by default (spec: avoid Folder1/Folder2 unless synthetic is chosen)", () => {
    const tree = buildFolderTree(createSeededSource(1), 5, 3, 2, "professional");
    for (const name of allNames(tree)) {
      expect(name).not.toMatch(/^Folder\d+$/);
    }
  });

  it("produces synthetic Folder1/Folder2/... names when namingStyle is 'synthetic'", () => {
    const tree = buildFolderTree(createSeededSource(1), 3, 2, 1, "synthetic");
    expect(tree.map((r) => r.name)).toEqual(["Folder1", "Folder2", "Folder3"]);
  });

  it("gives every root folder exactly subFoldersPerFolder children when nestedLevels >= 2", () => {
    const tree = buildFolderTree(createSeededSource(1), 4, 3, 2, "professional");
    for (const root of tree) expect(root.children).toHaveLength(3);
  });

  it("respects nestedLevels === 1 by never generating child folders", () => {
    const tree = buildFolderTree(createSeededSource(1), 4, 3, 1, "professional");
    for (const root of tree) expect(root.children).toEqual([]);
  });

  it("nests subFoldersPerFolder children at every level, down to nestedLevels deep", () => {
    const tree = buildFolderTree(createSeededSource(1), 2, 2, 3, "professional");
    for (const root of tree) {
      expect(root.children).toHaveLength(2);
      for (const child of root.children) {
        expect(child.children).toHaveLength(2);
        for (const grandchild of child.children) expect(grandchild.children).toEqual([]);
      }
    }
  });

  it("matches the user's own worked example: 20 total folders + 3 sub-folders each, one nested level", () => {
    const tree = buildFolderTree(createSeededSource(1), 20, 3, 2, "professional");
    expect(tree).toHaveLength(20);
    for (const root of tree) expect(root.children).toHaveLength(3);
    expect(countFolderNodes(tree)).toBe(20 + 20 * 3);
  });

  it("is deterministic for the same seed and inputs", () => {
    const a = buildFolderTree(createSeededSource(55), 6, 2, 3, "professional");
    const b = buildFolderTree(createSeededSource(55), 6, 2, 3, "professional");
    expect(a).toEqual(b);
  });

  it("never produces two folders with the same name anywhere in the tree, even once the name pool is exhausted", () => {
    const tree = buildFolderTree(createSeededSource(1), 20, 3, 2, "professional"); // 80 nodes, well past the ~70-name pool
    const names = allNames(tree);
    expect(new Set(names).size).toBe(names.length);
  });

  it("countFolderNodes counts every root, sub-folder, and nested folder", () => {
    const tree = buildFolderTree(createSeededSource(1), 3, 2, 3, "professional");
    expect(countFolderNodes(tree)).toBe(allNames(tree).length);
  });
});
