# PLINK GUI SNP Search Tool
A GUI to conduct SNP searches using Binary PED (".bed", ".bim", and ".fam") biobank files via Shaun Purcell's PLINK

This programme requires that you have PLINK, which can be downloaded here:
 - http://zzz.bwh.harvard.edu/plink/download.shtml
 
You should download this programme and place it in the same directory as your PLINK executable.

This programme requires that you place the binary files that you will use in the same directory as PLINK.

# How to use this programme - a brief overview
 - If you are only interested in searching for <b>one</b> SNP, simply type that SNP in and continue. You will receive a set of binary files derived from your source files that only contain the searched SNP.
 - If you have a list of <b>multiple</b> SNPs to search, and you would like these values included <b>cumulatively</b> in a new set of binary files, type "list" instead of an SNP. You will then be allowed to select a ".txt" file that contains all of your SNPs.
   - Your ".txt" file should be oriented somewhat like this:
   
     For example, I have "snps.txt" with the following contents:
     
         rs0001
         rs0002
         rs0003
         ...
         (Add as many as you'd like)
         
     I would select "snps.txt" for my SNP source file.
   - PLINK will be called to generate a single set of binary files with all listed SNPs.
