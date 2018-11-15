/*
Assumptions....
1. NO GLOBAL VARIABLES.
2. All Inserted Values are distinct.
3. Insert works even if the Tree is empty (Inserted node is made the root).
*/
#include <stdlib.h>
#include <stdio.h>
#define FOUND 1
#define NOTFOUND -9999
//Finding Max/Min.
typedef struct btree{
    int data;
    struct btree *left;
    struct btree *right;
}BST;
typedef struct node{
    BST *data;
    struct node *link;
}Node;
typedef struct stack{
    Node *top;
}Stack;

typedef struct queue{
    Node *front;
    Node *rear;
}Queue;
BST *getMaxNode(BST *tree){
    if(tree==NULL)
        return NULL;
    if(tree->right==NULL)
        return tree;
    return getMaxNode(tree->right);
}
BST *getMinNode(BST *tree){
    if(tree==NULL)
        return NULL;;
    if(tree->left==NULL)
        return tree;
    return getMinNode(tree->left);
}
//Searching.
int search(BST *tree,int S){
    if(tree==NULL)
        return NOTFOUND;
    if(S == tree->data)
        return FOUND;
    if(S < tree->data)
        return search(tree->left,S);
    return search(tree->right,S);
}
//Insert.
BST *insert(BST *tree,int v){
    if(tree==NULL){
        tree=(BST *)malloc(sizeof(BST));
        tree->data=v;
        tree->left=tree->right=NULL;
        printf("Inserted\n");
    }
    else
    if(v < tree->data)
        tree->left=insert(tree->left,v);
    else
    if(v > tree->data)
        tree->right=insert(tree->right,v);
    return tree;
}
//Delete.
BST *del(BST *tree,int v){
    if(tree==NULL)
        return NULL;
    if(v == tree->data){
        //Two children.
        if(tree->left!=NULL && tree->right != NULL){
            //In-order Successor.
            BST *suc=getMinNode(tree->right);
            tree->data=suc->data;
            tree->right=del(tree->right,tree->data);
            //In-order Predecessor.
            //BST *pre=getMaxNode(tree->left);
            //tree->data=pre->data;
            //tree->left=del(tree->left,pre->data);
        }
        //1 or No Child.
        else{
            BST *temp=tree;
            if(tree->left==NULL)
                tree=tree->right;
            else
            if(tree->right==NULL)
                tree=tree->left;
            free(temp);
        }
    }
    else
    if(v < tree->data)
        tree->left=del(tree->left,v);
    else
        tree->right=del(tree->right,v);
    return tree;
}

//Traversals.
void preorder(BST *tree){
    if(tree!=NULL){
        printf("%d ",tree->data);
        preorder(tree->left);
        preorder(tree->right);
    }
}
void inorder(BST *tree){
    if(tree!=NULL){
        inorder(tree->left);
        printf("%d ",tree->data);
        inorder(tree->right);
    }
}
void postorder(BST *tree){
    if(tree!=NULL){
        postorder(tree->left);
        postorder(tree->right);
        printf("%d ",tree->data);
    }
}
void iterative_preorder(BST *tree){
    if(tree==NULL)
        return;
    Stack *stack=(Stack *)malloc(sizeof(Stack));
    stack->top=NULL;//Imp in this particular implementation.
    push(stack,tree);
    while(!isEmpty(stack)){
        BST *t=pop(stack);
        printf("%d ",t->data);
        if(t->right!=NULL)
            push(stack,t->right);
        if(t->left!=NULL)
            push(stack,t->left);
    }
    printf("\n");
    free(stack);
}
void iterative_postorder(BST *tree){
    if(tree==NULL)
        return;
    Stack *stack=(Stack *)malloc(sizeof(Stack));
    stack->top=NULL;//Imp in this particular implementation.
    Stack *post=(Stack *)malloc(sizeof(Stack));
    post->top=NULL;
    push(stack,tree);
    while(!isEmpty(stack)){
        BST *t=pop(stack);
        push(post,t);
        if(t->left!=NULL)
            push(stack,t->left);
        if(t->right!=NULL)
            push(stack,t->right);
    }
    while(!isEmpty(post)){
        BST *t=pop(post);
        printf("%d ",t->data);
    }
    printf("\n");
    free(stack);
    free(post);
}
void iterative_preorder2(BST *tree){
    if(tree==NULL)
        return;
    Stack *stack=(Stack *)malloc(sizeof(Stack));
    stack->top=NULL;
    do{
        //Go left till childless.
        while(tree!=NULL){
            printf("%d ",tree->data);
            push(stack,tree);
            tree=tree->left;
        }
        if(isEmpty(stack))
            break;
        //Left subtree complete, so Pop.
        tree=pop(stack);
        //Go for the right subtree.
        tree=tree->right;
    }while(1);
    printf("\n");
}
void iterative_inorder2(BST *tree){
    if(tree==NULL)
        return;
    Stack *stack=(Stack *)malloc(sizeof(Stack));
    stack->top=NULL;
    do{
        //Go left till childless.
        while(tree!=NULL){
            push(stack,tree);
            tree=tree->left;
        }
        if(isEmpty(stack))
            break;
        //Left subtree complete, so Pop.
        tree=pop(stack);
        //Once left is visited then display.
        printf("%d ",tree->data);
        //Go for the right subtree.
        tree=tree->right;
    }while(1);
    printf("\n");
}
void iterative_postorder2(BST *tree){
    if(tree==NULL)
        return;
    Stack *stack=(Stack *)malloc(sizeof(Stack));
    stack->top=NULL;
    BST *prev;
    do{
        //Go left till childless.
        while(tree!=NULL){
            push(stack,tree);
            tree=tree->left;
        }
        while(tree==NULL && !isEmpty(stack)){
            tree=top(stack);
            if(tree->right==NULL || tree->right==prev){
                printf("%d ",tree->data);
                pop(stack);
                prev=tree;
                tree=NULL;
            }
            else
                tree=tree->right;
        }
        if(isEmpty(stack))
            break;
    }while(1);
    printf("\n");
}
void iterative_levelorder(BST *tree){
    if(tree==NULL)
        return;
    BST *temp;
    Queue *queue=(Queue *)malloc(sizeof(Queue));
    queue->front=queue->rear=NULL;
    enqueue(queue,tree);
    while(!isEmptyQ(queue)){
        temp=dequeue(queue);
        printf("%d ",temp->data);
        if(temp->left!=NULL)
            enqueue(queue,temp->left);
        if(temp->right!=NULL)
            enqueue(queue,temp->right);
    }
    printf("\n");
    free(queue);
}
