# Enter your code here. Read input from STDIN. Print output to STDOUT
s = input()
lst = []
count=1
i = 0
while i < len(s)-1:
    if s[i]==s[i+1]:
        count+=1
    else:
        lst.append([count,int(s[i])])
        count = 1
    i+=1
lst.append([count,int(s[i])])
lst = [tuple(i) for i in lst]
for i in lst:
    print (i,end=' ')