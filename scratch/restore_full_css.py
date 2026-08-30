import re

with open("scratch/orig_style.css", "r", encoding="utf-8") as f:
    orig_style = f.read()

# Custom styles we added for components and badges
new_css_additions = '''
        /* Component Badges on Subtask Checklist */
        .badge-tag-comp {
            display: inline-flex;
            align-items: center;
            font-size: 0.65rem;
            font-weight: 700;
            padding: 2px 7px;
            border-radius: 4px;
            letter-spacing: 0.04em;
            font-family: 'JetBrains Mono', monospace;
            text-transform: uppercase;
            white-space: nowrap;
        }
        .badge-tag-act {
            background: rgba(245, 158, 11, 0.15);
            color: #fbbf24;
            border: 1px solid rgba(245, 158, 11, 0.35);
        }
        .badge-tag-inst {
            background: rgba(6, 182, 212, 0.15);
            color: #38bdf8;
            border: 1px solid rgba(6, 182, 212, 0.35);
        }
        .badge-tag-elec {
            background: rgba(139, 92, 246, 0.15);
            color: #a78bfa;
            border: 1px solid rgba(139, 92, 246, 0.35);
        }

        /* Mode Selector Buttons for Add Subtask */
        .comp-mode-btn {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            color: var(--text-muted);
            padding: 5px 12px;
            border-radius: 6px;
            font-size: 0.78rem;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s ease;
            white-space: nowrap;
        }
        .comp-mode-btn:hover {
            border-color: var(--primary);
            color: var(--primary);
            background: rgba(99, 102, 241, 0.08);
        }
        .comp-mode-btn.active {
            background: var(--primary);
            color: #090d16;
            border-color: var(--primary);
            font-weight: 700;
            box-shadow: 0 0 8px rgba(99, 102, 241, 0.35);
        }

        /* Clean Checklist Item Layout */
        .checklist-item-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            width: 100%;
            min-height: 20px;
            margin-bottom: 2px;
        }
        .header-left {
            display: flex;
            align-items: center;
            gap: 6px;
            flex-wrap: wrap;
        }
        .header-right {
            display: flex;
            align-items: center;
            gap: 8px;
            margin-left: auto;
        }
        .btn-del-subtask-cross {
            background: transparent;
            border: none;
            color: var(--text-muted);
            font-size: 1.2rem;
            line-height: 1;
            cursor: pointer;
            padding: 0 4px;
            border-radius: 4px;
            transition: all 0.2s;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            opacity: 0.5;
        }
        .checklist-item:hover .btn-del-subtask-cross {
            opacity: 0.9;
        }
        .btn-del-subtask-cross:hover {
            color: #f43f5e !important;
            background: rgba(244, 63, 94, 0.15);
        }
'''

# Insert new additions before </style>
merged_style = orig_style.replace("</style>", new_css_additions + "\n    </style>")

with open("server.py", "r", encoding="utf-8") as f:
    server_code = f.read()

# Replace <style>...</style> in server.py
start_idx = server_code.find("<style>")
end_idx = server_code.find("</style>") + 8

new_server_code = server_code[:start_idx] + merged_style + server_code[end_idx:]

with open("server.py", "w", encoding="utf-8") as f:
    f.write(new_server_code)

print(f"Successfully restored full CSS! Old size: {end_idx - start_idx}, New size: {len(merged_style)}")
