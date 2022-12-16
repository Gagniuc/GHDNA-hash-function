# GHDNA hash function


A cryptographic hash function is a one-way function which provides a unique signature for any variable length sequence. Thus, any input leads to a unique output signature of a constant length. GHDNA is an experimental hash function that takes a DNA sequence as an argument and provides a unique signature in the output. The signature provided by the function is a constant length sequence. This project contains a series of independent applications that use the GHDNA hashing function. Some of them are used for testing, and others are used as demo applications.


| Input  | output |
| ------------- | ------------- |
| TATTCGGATCACGGACGTACC  | 74499889294608  |
| TATTCGGATCACGGACGTACA  | 40651499483769  |
| ATTCGGATCACGGACGTACC   | 44170498343430  |
| TTCGGATCACGGACGTACC    | 98964487625810  |
| ATCACGGAC              | 59451027176382  |
| GGGGGGGGGGGGGGGGGGGGG  | 84454969973368  |
| ACTT                   | 60847592169125  |



The main quality that a cryptographic hashing function must have is to evenly distribute the hashing keys over the domain range. This distribution is determined using the <kbd>GHDNA Domain test</kbd> application.


This experimental hash function also uses a new algorithm called Dynamic Block Allocation (DBA), which can be found [[here](https://github.com/Gagniuc/Dynamic-Block-Allocation-algorithm)]

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
