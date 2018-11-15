#include <stdio.h>
#include <stdlib.h>
typedef struct btree{
    int data;
    struct btree *left;
    struct btree *right;
}BST;
BST *create(BST *tree){
    if(tree!=NULL){
        printf("Tree already created. Remove it first to create a new one!\n");
        return tree;
    }
    printf("Enter elements one by one and -9999 to stop...\n");
    while(1){
        int ele;
        scanf("%d",&ele);
        if(ele==-9999)
            break;
        tree=insert(tree,ele);
    }
    return tree;
}
int main()
{
    BST *tree=NULL;
    do{
        printf("---------------MENU---------------\n");
        printf("1. Create a BST.\n");
        printf("2. Search for an element.\n");
        printf("3. Insert.\n");
        printf("4. Delete.\n");
        printf("5. Traverse (PRE-ORDER).\n");
        printf("6. Traverse (IN-ORDER).\n");
        printf("7. Traverse (POST-ORDER).\n");
        printf("8. Traverse (LEVEL-ORDER).\n");
        printf("9. Find the maximum element.\n");
        printf("10. Find the minimum element.\n");
        printf("11. Display the BST.\n");
        printf("12. Exit.\n");
        printf("----------------------------------\n");
        int ch,v,S,r;
        BST *max;
        BST *min;
        scanf("%d",&ch);
        switch(ch){
            case 1 :
                        tree=create(tree);
                        break;
            case 2 :
                        printf("Enter the element to be searched : ");
                        scanf("%d",&S);
                        r=search(tree,S);
                        if(r!=-9999)
                            printf("Element Found.\n");
                        else
                            printf("Element NOT Found.\n");
                        break;
            case 3 :
                        printf("Enter the element to be inserted : ");
                        scanf("%d",&v);
                        tree=insert(tree,v);
                        break;
                        break;
            case 4 :
                        printf("Enter the element to be deleted : ");
                        scanf("%d",&v);
                        tree=del(tree,v);
                        break;
            case 5 :
                        iterative_preorder2(tree);
                        printf("\n");
                        break;
            case 6 :
                        iterative_inorder2(tree);
                        printf("\n");
                        break;
            case 7 :
                        iterative_postorder2(tree);
                        printf("\n");
                        break;
            case 8 :
                        iterative_levelorder(tree);
                        printf("\n");
                        break;
            case 9 :
                        max=getMaxNode(tree);
                        if(max!=NULL)
                            printf("Max = %d\n",max->data);
                        else
                            printf("No BST Found.\n");
                        break;
            case 10 :
                        min=getMinNode(tree);
                        if(min!=NULL)
                            printf("Min = %d\n",min->data);
                        else
                            printf("No BST Found.\n");
                        break;
            case 11 :
                        print_ascii_tree(tree);
                        break;
            case 12 :
                        exit(0);
            default :
                        printf("Wrong Choice!");
        }
    }while(1);
    return 0;
}
