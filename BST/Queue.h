#ifndef QUEUE_H_INCLUDED
#define QUEUE_H_INCLUDED
extern BST;
extern Node;
typedef struct queue{
    Node *front;
    Node *rear;
}Queue;
int isEmptyQ(Queue *);
void ins(Queue *,BST *);
BST *rem(Queue *);
void delete_queue(Queue *);
#endif // QUEUE_H_INCLUDED
