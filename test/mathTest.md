template: Martin Template.pptx
baseTextSize: 20

# Math Test
Inline and block maths

### Inline Maths

Inline maths uses GitHub's second inline form: a dollar, a backtick, the LaTeX, a
backtick, a dollar. As with a block, nothing happens unless a stylesheet is
available - either beside md2pptx as mml2omml.xsl or named by mathxsl.

* The Euler identity $`e^{i\pi}+1=0`$ sits in the middle of a sentence
* Multiplication $`a*b`$ and a power $`x**2`$ keep their asterisks
* A bracketed fraction $`\left[\frac{a}{b}\right]`$ keeps its brackets
* Backslash escapes survive: $`f\_g`$ and $`\#\{x\}`$
* **Bold on either side of $`\mathcal{L}_{\text{patch}}`$ stays bold**
* *Italic either side of $`\lambda`$ stays italic*
* A sum with limits: $`\sum_{i=1}^{n} x_i^2`$
* A matrix: $`\begin{bmatrix}1 & 2 \\ 3 & 4\end{bmatrix}`$
* A bare dollar is not a delimiter, so costs $5 to $10 is left alone
* Neither is one in a link: [$`x`$](https://example.com) stays as source

### Inline Maths In A Table

A row is split on "|" before the text is parsed, so a formula in a cell has to
spell the bars as \vert or \mid.

| Symbol | Meaning |
|--------|---------|
| $`\left\vert x \right\vert`$ | absolute value |
| $`P(A \mid B)`$ | conditional probability |

### Block Maths

```math
\mathcal{L} = \mathcal{L}_{\text{patch}} + \lambda \mathcal{L}_{\text{glob}}
```

```math 1.2.4
\frac{-b \pm \sqrt{b^2-4ac}}{2a}
```
