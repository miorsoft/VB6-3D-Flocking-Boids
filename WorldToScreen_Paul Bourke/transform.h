#define DTOR 0.01745329252
#define EPSILON 0.001
#define PERSPECTIVE 0
#define ORTHOGRAPHIC 1
#define ABS(x) ((x) < 0 ? -(x) : (x))

/* Point in screen "window" space */
typedef struct {
   int h,v;
} Point;

/* Point in 3 space */
typedef struct {
   double x,y,z;
} XYZ;

/* Camera definition */
typedef struct {
   XYZ from;
   XYZ to;
   XYZ up;
   double angleh,anglev;
   double zoom;
   double front,back;
   short projection;
} CAMERA;

/* Screen definition */
typedef struct {
   Point center;
   Point size;
} SCREEN;

