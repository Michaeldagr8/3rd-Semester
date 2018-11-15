#include <stdio.h>
#include <stdlib.h>
typedef unsigned long long ull;
typedef long long ll;
extern ull comparisons;
void shuffle(ll *arr,ull n){
    ull i,r;
    for(i=0;i<n;i++){
        r=rand()%n;
        if(r!=i){
            arr[r]=arr[r]^arr[i];
            arr[i]=arr[r]^arr[i];
            arr[r]=arr[r]^arr[i];
        }
    }
}
ll *generate_sorted(ull n){
    ull i;
    ll *arr=(ll *)calloc((n+1),sizeof(ll));
    for(i=0;i<n;i++)
        arr[i]=i+1;
    return arr;
}
ll *generate_random(ull n){
    ll *arr=generate_sorted(n);
    if(arr==NULL){
        printf("Out of memory.\n");
        exit(0);
    }
    shuffle(arr,n);
    return arr;
}
int main()
{
    srand(time(NULL));
    ull n,pos;
    printf("Enter the size : ");
    scanf("%llu",&n);
    printf("Positions starting from 1....\n");
    ll *arr;
    arr=generate_random(n);
    //Linear Search.
    printf("Linear Search....\n");
    pos=linear_search(arr,n,arr[n-1]);
    if(pos!=-1)
        printf("No. found at position %llu.\n",pos);
    else
        printf("No. not found.");
    printf("No. of comparisons = %llu.\n",comparisons);
    printf("---------------------------------\n");
    //Linear Search with Sentinel.
    printf("Linear Search with Sentinel....\n");
    pos=linear_search_sentinel(arr,n,arr[n-1]);
    if(pos!=-1)
        printf("No. found at position %llu.\n",pos);
    else
        printf("No. not found.");
    printf("No. of comparisons = %llu.\n",comparisons);
    printf("---------------------------------\n");
    free(arr);
    arr=generate_sorted(n);
    //Binary Search.
    printf("Binary Search....\n");
    pos=binary_search(arr,n,arr[n-1]);
    if(pos!=-1)
        printf("No. found at position %llu.\n",pos);
    else
        printf("No. not found.");
    printf("No. of comparisons = %llu.\n",comparisons);
    printf("---------------------------------\n");
    //Interpolation Search.
    printf("Interpolation Search....\n");
    pos=interpolation_search(arr,n,arr[n-1]);
    if(pos!=-1)
        printf("No. found at position %llu.\n",pos);
    else
        printf("No. not found.");
    printf("No. of comparisons = %llu.\n",comparisons);
    printf("---------------------------------\n");
    //Interpolation Search (Worst Case).
    printf("Interpolation Search (Worst Case)....\n");
    arr[n]=arr[n-1]*arr[n-1];
    pos=interpolation_search(arr,n+1,arr[n-1]);
    if(pos!=-1)
        printf("No. found at position %llu.\n",pos);
    else
        printf("No. not found.");
    printf("No. of comparisons = %llu.\n",comparisons);
    printf("---------------------------------\n");
    free(arr);
    return 0;
}
