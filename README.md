# VBA-Sparse-Table
🚀 Optimizing Range Queries in Excel Using VBA

Given N numbers in the range from 1 to 10,000, we often want to quickly compute values like the minimum within a subrange [L, R] for each query. While Excel already supports this functionality, its time complexity is O(NQ) — which works fine for small datasets.

But what if we need to perform hundreds of thousands of queries on massive datasets?

💡 I developed a VBA macro that reduces the complexity to O(N log N + Q log N), allowing fast and efficient queries on large datasets directly from Excel.

Not only can this macro find minimum values — with small tweaks, it can also compute:

🧮 Sum

📈 Maximum

📉 Average

✖️ Product

The only requirement? The operation must be commutative.

This is a great example of combining algorithmic thinking with Excel automation to solve real-world data problems at scale.
