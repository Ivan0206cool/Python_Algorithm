# Algorithm 
# Binary Search 

def array(l:int,r:int,arr:list)->list:
   emp=[]
   for i in range(l,r+1):
      emp.append(arr[i])
   return emp

def Binary_Search(li:list,n:int)->int:
   l=0
   r=len(li)-1
   count=0
   while l<=r:
      mid=(l+r)//2
      if li[mid]==n:
         count+=1
         print(str(array(l,r,li))+" mid:"+str(li[mid])+",count:"+str(count))
         break
      print(str(array(l,r,li))+" mid:"+str(li[mid])+" count:"+str(count+1))
      if li[mid]>n:
         r=mid-1
      else:
         l=mid+1
      count+=1
   if li[mid] !=n:
      print("Not Find")       

Binary_Search([1,2,3,4,5,6,7,8,9],1)

e={}
for i in [2,3,3]:
   r=e.get(i,0)
   r+=1
   e[i]=r

print(e)