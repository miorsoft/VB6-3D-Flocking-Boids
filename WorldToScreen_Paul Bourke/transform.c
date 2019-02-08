#include "stdio.h"
#include "stdlib.h"
#include "math.h"
#include "transform.h"

/* Prototypes */
int  Trans_Initialise(CAMERA,SCREEN);
void Trans_World2Eye(XYZ,XYZ *,CAMERA);
int  Trans_ClipEye(XYZ *,XYZ *,CAMERA);
void Trans_Eye2Norm(XYZ,XYZ *,CAMERA);
int  Trans_ClipNorm(XYZ *,XYZ *);
void Trans_Norm2Screen(XYZ,Point *,SCREEN);
int  Trans_Point(XYZ,Point *,SCREEN,CAMERA);
int  Trans_Line(XYZ,XYZ,Point *,Point *,SCREEN,CAMERA);
void Normalise(XYZ *);
void CrossProduct(XYZ,XYZ,XYZ *);
int  EqualVertex(XYZ,XYZ);

/* Static globals */
double tanthetah,tanthetav;
XYZ basisa,basisb,basisc;

/*
   Trans_Init() initialises various variables and performs checks
   on some of the camera and screen parameters. It is left up to a
   particular implementation to handle the different error conditions.
   It should be called whenever the screen or camera variables change.
*/
int Trans_Initialise(camera,screen)
CAMERA camera;
SCREEN screen;
{
   XYZ origin = {0.0,0.0,0.0};

   /* Is the camera position and view vector coincident ? */
   if (EqualVertex(camera.to,camera.from)) {
      return(FALSE);
   }

   /* Is there a legal camera up vector ? */
   if (EqualVertex(camera.up,origin)) {
      return(FALSE);
   }
   
   basisb.x = camera.to.x - camera.from.x;
   basisb.y = camera.to.y - camera.from.y;
   basisb.z = camera.to.z - camera.from.z;
   Normalise(&basisb);
   
   CrossProduct(camera.up,basisb,&basisa);
   Normalise(&basisa);

   /* Are the up vector and view direction colinear */
   if (EqualVertex(basisa,origin)) {
      return(FALSE);
   }
   
   CrossProduct(basisb,basisa,&basisc);
   
   /* Do we have legal camera apertures ? */
   if (camera.angleh < EPSILON || camera.anglev < EPSILON) {
      return(FALSE);
   }
   
   /* Calculate camera aperture statics, note: angles in degrees */
   tanthetah = tan(camera.angleh * DTOR / 2);
   tanthetav = tan(camera.anglev * DTOR / 2);
   
   /* Do we have a legal camera zoom ? */
   if (camera.zoom < EPSILON) {
      return(FALSE);
   }
   
   /* Are the clipping planes legal ? */
   if (camera.front < 0 || camera.back < 0 || camera.back <= camera.front) {
      return(FALSE);
   }
   
   return(TRUE);
}

/*
   Take a point in world coordinates and transform it to
   a point in the eye coordinate system.
*/
void Trans_World2Eye(w,e,camera)
XYZ w;
XYZ *e;
CAMERA camera;
{
   /* Translate world so that the camera is at the origin */
   w.x -= camera.from.x;
   w.y -= camera.from.y;
   w.z -= camera.from.z;

   /* Convert to eye coordinates using basis vectors */
   e->x = w.x * basisa.x + w.y * basisa.y + w.z * basisa.z;
   e->y = w.x * basisb.x + w.y * basisb.y + w.z * basisb.z;
   e->z = w.x * basisc.x + w.y * basisc.y + w.z * basisc.z;
}

/*
   Clip a line segment in eye coordinates to the camera front
   and back clipping planes. Return FALSE if the line segment
   is entirely before or after the clipping planes.
*/
int Trans_ClipEye(e1,e2,camera)
XYZ *e1,*e2;
CAMERA camera;
{
   double mu;

   /* Is the vector totally in front of the front cutting plane ? */
   if (e1->y <= camera.front && e2->y <= camera.front)
      return(FALSE);
   
   /* Is the vector totally behind the back cutting plane ? */
   if (e1->y >= camera.back && e2->y >= camera.back)
      return(FALSE);
   
   /* Is the vector partly in front of the front cutting plane ? */
   if ((e1->y < camera.front && e2->y > camera.front) || 
      (e1->y > camera.front && e2->y < camera.front)) {
      mu = (camera.front - e1->y) / (e2->y - e1->y);
      if (e1->y < camera.front) {
         e1->x = e1->x + mu * (e2->x - e1->x);
         e1->z = e1->z + mu * (e2->z - e1->z);
         e1->y = camera.front;
      } else {
         e2->x = e1->x + mu * (e2->x - e1->x);
         e2->z = e1->z + mu * (e2->z - e1->z);
         e2->y = camera.front;
      }
   }

   /* Is the vector partly behind the back cutting plane ? */
   if ((e1->y < camera.back && e2->y > camera.back) || 
      (e1->y > camera.back && e2->y < camera.back)) {
      mu = (camera.back - e1->y) / (e2->y - e1->y);
      if (e1->y < camera.back) {
         e2->x = e1->x + mu * (e2->x - e1->x);
         e2->z = e1->z + mu * (e2->z - e1->z);
         e2->y = camera.back;
      } else {
         e1->x = e1->x + mu * (e2->x - e1->x);
         e1->z = e1->z + mu * (e2->z - e1->z);
         e1->y = camera.back;
      }
   }
   
   return(TRUE);
}

/*
   Take a vector in eye coordinates and transform it into
   normalised coordinates for a perspective view. No normalisation
   is performed for an orthographic projection. Note that although
   the y component of the normalised vector is copied from the eye
   coordinate system, it is generally no longer needed. It can
   however still be used externally for vector sorting.
*/
void Trans_Eye2Norm(e,n,camera)
XYZ e;
XYZ *n;
CAMERA camera;
{
	double d;
	
   if (camera.projection == PERSPECTIVE) {
   	d = camera.zoom / e.y;
      n->x = d * e.x / tanthetah;
      n->y = e.y;;
      n->z = d * e.z / tanthetav;
   } else {
      n->x = camera.zoom * e.x / tanthetah;
      n->y = e.y;
      n->z = camera.zoom * e.z / tanthetav;
   }
}

/* 
   Clip a line segment to the normalised coordinate +- square.
   The y component is not touched.
*/
int Trans_ClipNorm(n1,n2)
XYZ *n1,*n2;
{
   double mu;

   /* Is the line segment totally right of x = 1 ? */
   if (n1->x >= 1 && n2->x >= 1)
      return(FALSE);

   /* Is the line segment totally left of x = -1 ? */
   if (n1->x <= -1 && n2->x <= -1)
      return(FALSE);
      
   /* Does the vector cross x = 1 ? */
   if ((n1->x > 1 && n2->x < 1) || (n1->x < 1 && n2->x > 1)) {
      mu = (1 - n1->x) / (n2->x - n1->x);
      if (n1->x < 1) {
         n2->z = n1->z + mu * (n2->z - n1->z);
         n2->x = 1;
      } else {
         n1->z = n1->z + mu * (n2->z - n1->z);
         n1->x = 1;
      }
   }
      
   /* Does the vector cross x = -1 ? */
   if ((n1->x < -1 && n2->x > -1) || (n1->x > -1 && n2->x < -1)) {
      mu = (-1 - n1->x) / (n2->x - n1->x);
      if (n1->x > -1) {
         n2->z = n1->z + mu * (n2->z - n1->z);
         n2->x = -1;
      } else {
         n1->z = n1->z + mu * (n2->z - n1->z);
         n1->x = -1;
      }
   }

   /* Is the line segment totally above z = 1 ? */
   if (n1->z >= 1 &&; n2->z >= 1)
      return(FALSE);

   /* Is the line segment totally below z = -1 ? */
   if (n1->z <= -1 && n2->z <= -1)
      return(FALSE);
      
   /* Does the vector cross z = 1 ? */
   if ((n1->z > 1 && n2->z < 1) || (n1->z < 1 && n2->z > 1)) {
      mu = (1 - n1->z) / (n2->z - n1->z);
      if (n1->z < 1) {
         n2->x = n1->x + mu * (n2->x - n1->x);
         n2->z = 1;
      } else {
         n1->x = n1->x + mu * (n2->x - n1->x);
         n1->z = 1;
      }
   }
      
   /* Does the vector cross z = -1 ? */
   if ((n1->z < -1 && n2->z > -1) || (n1->z > -1 && n2->z < -1)) {
      mu = (-1 - n1->z) / (n2->z - n1->z);
      if (n1->z > -1) {
         n2->x = n1->x + mu * (n2->x - n1->x);
         n2->z = -1;
      } else {
         n1->x = n1->x + mu * (n2->x - n1->x);
         n1->z = -1;
      }
   }
   
   return(TRUE);
}

/*
   Take a vector in normalised Coordinates and transform it into
   screen coordinates.
*/
void Trans_Norm2Screen(norm,projected,screen)
XYZ norm;
Point *projected;
SCREEN screen;
{
   projected->h = screen.center.h - screen.size.h * norm.x / 2;
   projected->v = screen.center.v - screen.size.v * norm.z / 2;
}

/* 
   Transform a point from world to screen coordinates. Return TRUE
   if the point is visible, the point in screen coordinates is p.
   Assumes Trans_Initialise() has been called
*/
int Trans_Point(w,p,screen,camera)
XYZ w;
Point *p;
SCREEN screen;
CAMERA camera;
{
   XYZ e,n;
   
   Trans_World2Eye(w,&e,camera);
   if (e.y >= camera.front && e.y <= camera.back) {
      Trans_Eye2Norm(e,&n,camera);
      if (n.x >= -1 && n.x <= 1 && n.z >= -1 && n.z <= 1) {
         Trans_Norm2Screen(n,p,screen);
         return(TRUE);
      }
   }
   return(FALSE);
}

/* 
   Transform and appropriately clip a line segment from
   world to screen coordinates. Return TRUE if something
   is visible and needs to be drawn, namely a line between
   screen coordinates p1 and p2.
   Assumes Trans_Initialise() has been called
*/
int Trans_Line(w1,w2,p1,p2,screen,camera)
XYZ w1,w2;
Point *p1,*p2;
SCREEN screen;
CAMERA camera;
{
   XYZ e1,e2,n1,n2;
   
   Trans_World2Eye(w1,&e1,camera);
   Trans_World2Eye(w2,&e2,camera);
   if (Trans_ClipEye(&e1,&e2,camera)) {
      Trans_Eye2Norm(e1,&n1,camera);
      Trans_Eye2Norm(e2,&n2,camera);
      if (Trans_ClipNorm(&n1,&n2)) {
         Trans_Norm2Screen(n1,p1,screen);
         Trans_Norm2Screen(n2,p2,screen);
         return(TRUE);
      }
   }
   return(FALSE);
}

/*
   Normalise a vector
*/
void Normalise(v)
XYZ *v;
{
   double length;
	
   length = sqrt(v->x * v->x + v->y * v->y + v->z * v->z);
   v->x /= length;
   v->y /= length;
   v->z /= length;
}

/*
   Cross product of two vectors, p3 = p1 x p2
*/
void CrossProduct(p1,p2,p3)
XYZ p1,p2,*p3;
{
   p3->x = p1.y * p2.z - p1.z * p2.y;
   p3->y = p1.z * p2.x - p1.x * p2.z;
   p3->z = p1.x * p2.y - p1.y * p2.x;
}

/*
   Test for coincidence of two vectors, TRUE if cooincident
*/
int EqualVertex(p1,p2)
XYZ p1,p2;
{
   if (ABS(p1.x - p2.x) > EPSILON)
      return(FALSE);
   if (ABS(p1.y - p2.y) > EPSILON)
      return(FALSE);
   if (ABS(p1.z - p2.z) > EPSILON)
      return(FALSE);
   return(TRUE);
}
