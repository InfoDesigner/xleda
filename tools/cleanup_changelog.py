# import re
# from pathlib import Path


# def cleanup_readme_changelog(readme_path: str | Path) -> None:
#     """
#         Removes all but the last Changelog entry and converts it to standard markdown

#     Args:
#         readme_path: Path to README.md file
        
#     Raises:
#         RuntimeError: If the cleanup fails
#     """
#     readme_path = Path(readme_path)
    
#     try:
#         # Read the file
#         with open(readme_path, 'r', encoding='utf-8') as f:
#             content = f.read()
#     except Exception as e:
#         raise RuntimeError(f"Failed to read README.md: {e}")
    
#     try:
#         # Find the Changelog section
#         changelog_match = re.search(r'## Changelog\n', content)
#         if not changelog_match:
#             raise RuntimeError("Changelog section not found in README.md")
        
#         changelog_start = changelog_match.end()
        
#         # Find all <details> blocks after Changelog header
#         # Pattern to match entire <details>...</details> blocks
#         details_pattern = r'<details>\s*\n\s*<summary>([^<]*)</summary>\s*\n(.*?)\n</details>'
        
#         details_blocks = list(re.finditer(details_pattern, content[changelog_start:], re.DOTALL))
        
#         if not details_blocks:
#             raise RuntimeError("No <details> blocks found in Changelog section")
        
#         if len(details_blocks) < 2:
#             raise RuntimeError(f"Expected at least 2 <details> blocks, found {len(details_blocks)}")
        
#         # Keep only the last block
#         last_block = details_blocks[-1]
        
#         # Extract summary text and content
#         summary_text = last_block.group(1).strip()
#         block_content = last_block.group(2).strip()
        
#         # Find the position of the last block in the original content
#         last_block_start = changelog_start + last_block.start()
#         last_block_end = changelog_start + last_block.end()
        
#         # Find the position of the first block
#         first_block_start = changelog_start + details_blocks[0].start()
#         first_block_end = changelog_start + details_blocks[-2].end() if len(details_blocks) > 1 else changelog_start + details_blocks[0].end()
        
#         # For multiple blocks, we need to remove all but the last
#         # Calculate the range of all blocks except the last
#         all_but_last_start = changelog_start + details_blocks[0].start()
#         all_but_last_end = changelog_start + details_blocks[-2].end()
        
#         # Remove all blocks except the last
#         content_before = content[:all_but_last_start]
#         content_after = content[all_but_last_end:]
        
#         # Now find and transform the last block in content_after
#         last_block_match = re.search(
#             r'<details>\s*\n\s*<summary>([^<]*)</summary>\s*\n(.*?)\n</details>',
#             content_after,
#             re.DOTALL
#         )
        
#         if not last_block_match:
#             raise RuntimeError("Failed to locate last <details> block for transformation")
        
#         summary_text = last_block_match.group(1).strip()
#         block_content = last_block_match.group(2).strip()
        
#         # Transform: create markdown version
#         markdown_block = f"### {summary_text}\n\n{block_content}\n"
        
#         # Replace the last block with the markdown version
#         transformed_content_after = content_after[:last_block_match.start()] + markdown_block + content_after[last_block_match.end():]
        
#         # Combine
#         final_content = content_before + transformed_content_after
        
#         # Write back to file
#         with open(readme_path, 'w', encoding='utf-8') as f:
#             f.write(final_content)
            
#     except RuntimeError:
#         raise
#     except Exception as e:
#         raise RuntimeError(f"Failed to process README.md: {e}")
    
#     print("README was successfully adjusted")


# if __name__ == "__main__":
#     readme_path = Path(__file__).parent.parent / "README.md"
#     cleanup_readme_changelog(readme_path)
