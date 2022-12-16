# GHDNA hash function


A cryptographic hash function is a one-way function which provides a unique signature for any variable length sequence. Thus, any input leads to a unique output signature of a constant length. GHDNA is an experimental hash function that takes a DNA sequence as an argument and provides a unique signature in the output. The signature provided by the function is a constant length sequence. This project contains a series of independent applications that use the GHDNA hashing function. Some of them are used for testing, and others are used as demo applications.


| Input  | output |
| ------------- | ------------- |
| TATTCGGATCACGGACGTACC  | 74499889294608  |
| TATTCGGATCACGGACGTACA  | 40651499483769  |
| ATTCGGATCACGGACGTACC   | 44170498343430  |
| TTCGGATCACGGACGTACC    | 98964487625810  |
| ATCACGGAC              | 59451027176382  |



A second important quality is that these hashing keys have associated as few imputations as possible from the set of imputations or the association of multiple imputations should be equally distributed on the key stack.

In other words, if a cryptographic hashing function produces 10 hashing keys and the total of possible imputations is 20 inputs, then it would be ideal for each cryptographic key to be associated with 2 imputations. This example demonstrates how interesting cryptographic functions are and how difficult these ideal situations can be obtained. In the case of GHDNA, the number of keys is 10^14 and the total number of possible arguments tends to infinity. But, from our reference system, in practice it uses a limited number of arguments, because our informational set is limited and guided by the environment.





The main quality that a cryptographic hashing function must have is to evenly distribute the hashing keys over the domain range. This distribution is determined using the <kbd>GHDNA Domain test</kbd> application.

The GHDNA and GHDNA DATA BLOCK applications represent the simple version that shows how the GHDNA function can be used directly.

The GHDNA Avalanche test application demonstrates how tiny changes in the input sequence can generate totally different and unpredictable hashing keys. Collisions are also tested, where it is checked if a hashing key is also associated with another previous imput.

The GHDNA Speed test application measures the processing time of the GHDNA function in order to be able to compare it with other cryptographic functions (please see the attached article). In terms of processing time, GHDNA is very fast compared to the existing ones, but this speed is a bit relative and may also be due to the lack of complexity when compared to the often used cryptographic hashing functions.

The GHDNA database engine application uses the GHDNA cryptographic function to perform a demonstration within a primitive database engine.


This experimental hash function also uses a new algorithm called Dynamic Block Allocation (DBA), which can be found [[here](https://github.com/Gagniuc/Dynamic-Block-Allocation-algorithm)]

Note: in the BASIC family of computer languages, the "^" character represents exponentiation.

<kbd><img src="https://github.com/Gagniuc/GHDNA-hash-function/blob/main/img/1.png?raw=true" /></kbd>
<kbd><img src="https://github.com/Gagniuc/GHDNA-hash-function/blob/main/img/2.png?raw=true" /></kbd>
<kbd><img src="https://github.com/Gagniuc/GHDNA-hash-function/blob/main/img/3.png?raw=true" /></kbd>
<kbd><img src="https://github.com/Gagniuc/GHDNA-hash-function/blob/main/img/4.png?raw=true" /></kbd>
<kbd><img src="https://github.com/Gagniuc/GHDNA-hash-function/blob/main/img/5.png?raw=true" /></kbd>
<kbd><img src="https://github.com/Gagniuc/GHDNA-hash-function/blob/main/img/6.png?raw=true" /></kbd>
<kbd><img src="https://github.com/Gagniuc/GHDNA-hash-function/blob/main/img/7.png?raw=true" /></kbd>
<kbd><img src="https://github.com/Gagniuc/GHDNA-hash-function/blob/main/img/8.png?raw=true" /></kbd>

# References

- <i>Paul A. Gagniuc and Constantin Ionescu-Tîrgovişte. GHDNA: a hash function for DNA segment-based aligments and motif search. Proc. Rom. Acad., Series B, 2014, 16(3), p. 155–167.</i>
- <i>P. Gagniuc and C Ionescu-Tirgoviste. Dynamic block allocation for biological sequences. Proc. Rom. Acad., Series B, 2013, 15(3), p. 233-240.</i> 
