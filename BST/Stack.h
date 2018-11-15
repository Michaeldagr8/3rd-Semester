#ifndef STACK_H_INCLUDED
#define STACK_H_INCLUDED
BST;
Node;
typedef struct stack{
    Node *top;
}Stack;
int isEmpty(Stack *);
void push(Stack *,BST *);
BST *pop(Stack *);
BST *top(Stack *);
void delete_stack(Stack *);
#endif // STACK_H_INCLUDED
