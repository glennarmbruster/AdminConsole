#!/usr/bin/env python3
"""Brace/paren balance check for .ps1 files, comment- and string-aware. Usage: balance-check.py <folder>"""
import sys,glob,os
def strip(s):
    out=[];i=0;n=len(s)
    while i<n:
        c=s[i]
        if s[i:i+2]=='<#':
            d=1;i+=2
            while i<n and d:
                if s[i:i+2]=='<#':d+=1;i+=2
                elif s[i:i+2]=='#>':d-=1;i+=2
                else:i+=1
            continue
        if c=='#':
            while i<n and s[i]!='\n':i+=1
            continue
        if c=="'":
            i+=1
            while i<n:
                if s[i]=="'":
                    if s[i+1:i+2]=="'":i+=2;continue
                    i+=1;break
                i+=1
            continue
        if c=='"':
            i+=1
            while i<n:
                if s[i]=='`':i+=2;continue
                if s[i]=='"':
                    if s[i+1:i+2]=='"':i+=2;continue
                    i+=1;break
                i+=1
            continue
        if s[i:i+2] in ('@"',"@'"):
            q=s[i+1];i+=2;end='\n'+q+'@';j=s.find(end,i);i=(j+len(end)) if j>=0 else n;continue
        out.append(c);i+=1
    return ''.join(out)
bad=0;files=sorted(glob.glob(os.path.join(sys.argv[1],'*.ps1')))
for f in files:
    t=strip(open(f,encoding='utf-8-sig').read())
    for o,c,name in (('{','}','brace'),('(',')','paren')):
        d=t.count(o)-t.count(c)
        if d: print(f"UNBALANCED {name} {d:+d}  {os.path.basename(f)}");bad+=1
print(f"checked {len(files)} files; {bad} problems")
