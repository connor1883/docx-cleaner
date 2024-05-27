Econ 204 // Spring 2024

2024-01-29

**Quantitative Variables in Finite Populations and Samples:**

**Notation, Definitions, and Facts**

**I. QUANTITATIVE VARIABLES IN POPULATIONS**

[Notation]{.underline}

A ***finite population*** consists of \\(N\\) elements, labelled \\(I =
1,\...,N\\), where \\(N\\) is a real number.

\\(X\\) is the name of a quantitative variable (QVAR).

For \\(I = 1,\...,N\\), \\({X_I}\\) is the value of \\(X\\) for elements
\\(I\\).

[Definition: The population mean of a QVAR]{.underline}

The ***population mean of a QVAR*** \\(X\\) is represented by the symbol
\\({\\mu \_X}\\) , and defined as

\\({\\mu \_X} = \\frac{{\\sum\\limits\_{I = 1}\^N {{X_I}} }}{N}\\)

[Definition: A deviation from the population mean of a QVAR for a given
element]{.underline}

A ***deviation from the population*** ***mean of a QVAR*** \\(X\\)
***for a given element*** is the difference between the value of \\(X\\)
for the given element and the population mean of \\(X.\\) The deviation
from the population mean of \\(X\\) for element \\(I\\)is represented by
the symbol \\({D\_{{X_I}}}\\):

\\({D\_{{X_I}}} = {X_I} - {\\mu \_X}\\)

[Fact]{.underline}

For any finite population and any QVAR \\(X\\), if you find the
deviation from the population mean of \\(X\\) for every element in the
population, and then add up all those deviations, you will find that the
sum is zero.

To put that more succinctly: for any finite population and any
quantitative variable \\(X\\), the sum of the deviations from the means
is always equal to 0.

This is true no matter how many elements there are in the population and
no matter what the values of \\(X\\) are for the elements.

More formally: Given a finite population consisting of any number
\\(N\\) of elements, and given any values
\\({X_1},\\,{X_2},\\,\...,\\,{X_N}\\) of a QVAR \\(X\\),

> \\(\\sum\\limits\_{I = 1}\^N {{D\_{{X_I}}}} = 0\\)

or equivalently,

> \\(\\sum\\limits\_{I = 1}\^N {\\left( {{X_I} - {\\mu \_X}} \\right)} =
> 0\\)

[Proof of this fact]{.underline}

We can prove this equality through several steps of algebra. Every step
is just a rearrangement of terms.

> \\\[\\begin{array}{l}\\sum\\limits\_{I = 1}\^N {{D\_{{X_I}}}} =
> \\sum\\limits\_{I = 1}\^N {\\left( {{X_I} - {\\mu \_X}} \\right)}
> \\\\\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\, = \\sum\\limits\_{I
> = 1}\^N {{X_I}} - \\sum\\limits\_{I = 1}\^N {{\\mu \_X}}
> \\\\\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\, = \\sum\\limits\_{I
> = 1}\^N {{X_I}} - N{\\mu
> \_X}\\\\\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\, = N\\left(
> {\\frac{{\\sum\\limits\_{I = 1}\^N {{X_I}} }}{N}} \\right) - N{\\mu
> \_X}\\\\\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\, = N{\\mu \_X} -
> N{\\mu \_X}\\\\\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\, =
> 0\\end{array}\\\]

[Definition: The population mean absolute deviation of a
QVAR]{.underline}

The ***population mean absolute deviation*** ***of a QVAR*** \\(X\\) is
represented by the symbol \\\[MA{D_X}\\\], and defined as

\\\[MA{D_X} = \\frac{{\\sum\\limits\_{I = 1}\^N {\\left\| {{X_I} - {\\mu
\_X}} \\right\|} }}{N}\\\]

[Definition: The population variance of a QVAR]{.underline}

The ***population variance*** ***of a QVAR*** \\(X\\) is represented by
the symbol \\\[\\sigma \_X\^2\\\], and defined as

\\\[\\sigma \_X\^2 = \\frac{{\\sum\\limits\_{I = 1}\^N {{{\\left(
{{X_I} - {\\mu \_X}} \\right)}\^2}} }}{N}\\\]

[Definition: The population standard deviation of a QVAR]{.underline}

The ***population standard deviation*** ***of a QVAR*** \\(X\\) is
represented by the symbol \\\[{\\sigma \_X}\\\], and defined as

\\\[{\\sigma \_X} = \\sqrt\[ + \]{{\\sigma \_X\^2}}\\\]

**II. QUANTITATIVE VARIABLES IN SAMPLES**

The notation we use for samples is different from the notation we use
for populations.

For some concepts, the definitions for samples are identical to the
definitions for populations, and the only differences are in the
notation.

For other concepts, the differences are not only in the notation---the
definitions for samples are similar, but not identical, to the
definitions for populations.

[Notation]{.underline}

A ***sample*** consists of \\(n\\) observations, labelled \\(i =
1,\...,n\\).

\\(X\\) is the name of a quantitative variable.

For \\(i = 1,\...,n\\), \\({x_i}\\) is the value of \\(X\\) for
observation \\(i\\).

[Definition: The sample mean of a QVAR]{.underline}

The ***sample mean of a QVAR*** \\(X\\) is represented by the symbol
\\(\\bar x\\) , and defined as

\\(\\bar x = \\frac{{\\sum\\limits\_{i = 1}\^n {{x_i}} }}{n}\\)

[Definition: A deviation from the sample mean of a QVAR for a given
observation]{.underline}

A ***deviation from the sample mean of a QVAR*** \\(X\\) ***for a given
observation*** is the difference between the value of \\(X\\) for the
given observation and the population mean of \\(X\\). The deviation from
the sample mean of \\(Q\\) for observation \\(i\\) is represented by the
symbol \\({d\_{{X_i}}}\\):

\\({d\_{{X_i}}} = {x_i} - \\bar x\\)

[Fact]{.underline}

For every sample and any QVAR \\(X\\), if you find the deviation from
the sample mean of \\(X\\) for every observation in the sample, and then
add up all those deviations, you will find that the sum is zero.

To put that more succinctly: for any sample and any quantitative
variable \\(X\\), the sum of the deviations from the means is always
equal to 0.

This is true no matter how many observations there are in the sample and
no matter what the values of \\(X\\) are for the observations.

More formally: Given a sample consisting of any number \\(n\\) of
observations, and given any values \\({x_1},\\,{x_2},\\,\...,\\,{x_n}\\)
of a QVAR \\(X\\),

> \\(\\sum\\limits\_{i = 1}\^n {{d\_{{X_i}}}} = 0\\)

[Proof of this fact]{.underline}

We can prove this equality through several steps of algebra. Every step
is just a rearrangement of terms.

> \\\[\\begin{array}{l}\\sum\\limits\_{i = 1}\^n {{d\_{{X_i}}}} =
> \\sum\\limits\_{i = 1}\^n {\\left( {{x_i} - \\bar x} \\right)}
> \\\\\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\, = \\sum\\limits\_{i
> = 1}\^n {{x_i}} - \\sum\\limits\_{i = 1}\^n {\\bar x}
> \\\\\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\, = \\sum\\limits\_{i
> = 1}\^n {{x_i}} - n\\bar
> x\\\\\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\, = n\\left(
> {\\frac{{\\sum\\limits\_{i = 1}\^n {{x_i}} }}{n}} \\right) - n\\bar
> x\\\\\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\, = n\\bar x - n\\bar
> x\\\\\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\,\\, = 0\\end{array}\\\]

[Definition: The sample mean absolute deviation of a QVAR]{.underline}

The ***sample mean absolute deviation*** ***of a QVAR*** \\(X\\) is
represented by the symbol \\\[ma{d_X}\\\], and defined as

\\\[ma{d_X} = \\frac{{\\sum\\limits\_{i = 1}\^n {\\left\| {{x_i} - \\bar
x} \\right\|} }}{{n - 1}}\\\]

[Definition: The sample variance of a QVAR]{.underline}

The ***sample variance*** ***of a QVAR*** \\(X\\) is represented by the
symbol \\\[s_X\^2\\\], and defined as

\\\[s_X\^2 = \\frac{{\\sum\\limits\_{i = 1}\^n {{{\\left( {{x_i} - \\bar
x} \\right)}\^2}} }}{{n - 1}}\\\]

[Definition: The sample standard deviation of a QVAR]{.underline}

The ***sample standard deviation*** ***of a QVAR*** \\(X\\) is
represented by the symbol \\\[{s_X}\\\], and defined as

\\\[{s_X} = \\sqrt\[ + \]{{s_X\^2}}\\\]