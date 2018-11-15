typedef unsigned long long ull;
typedef long long ll;
ull comparisons;
ull linear_search(ll *arr,ull n,ll S){
    comparisons=1;
    ull i;
    for(i=0;i<n;i++){
        comparisons++;//If  condition.
        if(arr[i]==S)
            return i+1;
        comparisons ++;//Loop condition.
    }
    return -1;
}
ull linear_search_sentinel(ll *arr,ull n,ll S){
    comparisons=1;
    arr[n]=S;
    ull i;
    for(i=0;arr[i]!=S;i++)
        comparisons++;//Loop condition.
    comparisons ++;//If condition.
    if(i!=n)
        return i+1;
    return -1;
}
ull binary_search(ll *arr,ull n,ll S){
    comparisons=1;
    ull sp=0,ep=n-1,mid;
    while(sp<=ep){
        mid=(sp+ep)/2;
        comparisons++;//If condition.
        if(arr[mid]==S)
            return mid+1;
        comparisons++;//If condition.
        if(arr[mid]<S)
            sp=mid+1;
        else
            ep=mid-1;
        comparisons++;//Loop condition.
    }
    return -1;
}
ull interpolation_search(ll *arr,ull n,ll S){
    comparisons=1;
    ull sp=0,ep=n-1,p;
    while(sp<=ep){
        p=sp+((ep-sp)*(S-arr[sp]+1))/(arr[ep]-arr[sp]+1);
        comparisons++;//If condition.
        if(arr[p]==S)
            return p+1;
        comparisons++;//If condition.
        if(arr[p]<S)
            sp=p+1;
        else
            ep=p-1;
        comparisons ++;//Loop condition.
    }
    return -1;
}
