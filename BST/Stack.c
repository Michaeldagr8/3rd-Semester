#define TRUE 1
#define FALSE 0
#include <stdlib.h>
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
int isEmpty(Stack *stack){
    if(stack->top==NULL)
        return TRUE;
    return FALSE;
}
void push(Stack *stack,BST *v){
    Node *newnode=(Node *)malloc(sizeof(Node));
    newnode->data=v;
    newnode->link=stack->top;
    stack->top=newnode;
}
BST *pop(Stack *stack){
    if(isEmpty(stack))
        return -9999;
    BST *ele=stack->top->data;
    Node *temp=stack->top;
    stack->top=stack->top->link;
    free(temp);
    return ele;
}
BST *top(Stack *stack){
    if(isEmpty(stack))
        return -9999;
    return stack->top->data;
}
void delete_stack(Stack *stack){
    while(!isEmpty(stack))
        pop(stack);
}
