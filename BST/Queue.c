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
typedef struct queue{
    Node *front;
    Node *rear;
}Queue;
int isEmptyQ(Queue *queue){
    if(queue->front==NULL)
        return TRUE;
    return FALSE;
}
void enqueue(Queue *queue,BST *tree){
    Node *newnode=(Node *)malloc(sizeof(Node));
    newnode->data=tree;
    newnode->link=NULL;
    if(isEmptyQ(queue)){
        queue->front=newnode;
        queue->rear=newnode;
    }
    else{
        queue->rear->link=newnode;
        queue->rear=newnode;
    }
}
BST *dequeue(Queue *queue){
    if(isEmptyQ(queue))
        return -9999;
    BST *ele=queue->front->data;
    Node *temp=queue->front;
    queue->front=queue->front->link;
    free(temp);
    if(queue->front==NULL)
        queue->rear=NULL;
    return ele;
}
void delete_queue(Queue *queue){
    while(!isEmptyQ(queue))
        dequeue(queue);
}
