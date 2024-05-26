# LogicGame

Build logic with XOR gates!

# Editing board

Use LMB to draw wire or change wire state. Use RMB to erase. Modifier keys can extend erasing range.

You can clear, import, export as text, export as image (from/to clipboard) board using buttons on menu bar.

# Logic

When an enpty cell is surrounded by 4 wires, it acts as a bridge.

When an enpty cell is surrounded by 3 wires, the middle one is pulled exclusive or of two other values. (Black as 0, red as 1)

If a wire is pulled 1 more than pulled 0, it becomes 1; 0 if pulled 0 more than 1; not changed if equally pulled to 0 and 1.

# Simulate

Press space to step. SDFGH to start simulation, S is slowest and H is fastest. A to stop simulation.
