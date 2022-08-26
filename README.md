# trtc
My attempt at The Raytracer Challenge by Jamis Buck. I typically work with legacy MS Office code so this was my attempt to push the VBA envelope. 
It is also an opportunity to use/learn class modules

-----------------------------------------------------------------------------
 Chapter 1. Tuples, Points, and Vectors
-----------------------------------------------------------------------------
 As defined in the book, a Tuple is an ordered list. We are working with 3D
 coordinates (points & vectors) so we define elements to capture X, Y, and Z.
 To distinguish between a point and a vector the book use W. W=1 for points and
 W=0 for vectors

 Since I don't have inheritence or public Function overloading, I am working
 purely with tuples and dont have a point and vector class

 Not all math operations are applicable for tuples, points, and vectors but
 this is not a library for the public and assume they will be used correctly.

 Uses newTuple, newPoint and newVector functions defined in the Helper
 module to create objects

-----------------------------------------------------------------------------
 Chapter 2. Drawing on a Canvas
-----------------------------------------------------------------------------
 Currently modeling the Color class as a tuple so we include some RGB
 reference (newColor in the Helper module)

 Modeled output to PPM as identified in this chapter

 Updated model to output to a worksheet using resized cells as pixels

 Updated model to output to BMP using code found on the WWW

 Updated model to capture WKS output to various image formats using code
 found on the WWW.  Captures the worksheet onto a chart object as and uses
 the chartobject.save feature to write the various formats

-----------------------------------------------------------------------------
 Chapter 3. Matrices
-----------------------------------------------------------------------------
 2x2, 3x3, and 4x4 matrix math

-----------------------------------------------------------------------------
 Chapter 4. Matrix Transformations
-----------------------------------------------------------------------------
 Build transformations for Translate, Scale, Rotate, Shear

 Chaining
 Note that the order of the multiplications is important! Matrix
 multiplication is associative, but not commutative. When it comes to
 matrices, A × B is not guaranteed to be the same as B × A.

 So, if you want a single matrix that rotates, and then scales, and then
 translates, you can multiply the translation matrix by the scaling matrix,
 and then by the rotation matrix. That is to say, you must concatenate the
 transformations in reverse order to have them applied in the order you want!

 teapot -->  C * (B * (A * teapot)) = (C * B * A) * teapot

-----------------------------------------------------------------------------
 Chapter 5. Ray-Sphere Intersections
-----------------------------------------------------------------------------
 Create Rays and logic for intersection with a unit sphere. Also presents
 logic showing transformation relationship between ray and sphere. Logic
 for tracking Ray-Sphere intersections and which are visible or "Hits".

 Results in initial logic to display a Sphere outline in a 2D plane

![images/Ch05.png](images/Ch05.png)


-----------------------------------------------------------------------------
 Chapter 6. Light and Shading
-----------------------------------------------------------------------------
 Implements Phong reflective model. Model creates and uses light and
 material objects. Light is modeled as a Pointlight

 Updated model to work with multiple light sources (Ch 7 sidebar). Interface
 needs to change, Collections can't be updated and light sources are replaced
 during some of the tests

 Updated model to handle Spotlight (Ch 17). Need to handle fade

![images/Ch06.png](images/Ch06.png)



# Chapter 7 
![images/Ch07.png](images/Ch07.png)

# Chapter 9 
![images/Ch09.png](images/Ch09.png)


# Chapter 11-1
![images/Ch11_1.png](images/Ch11_1.png)

# Chapter 11-2
![images/Ch11_2.png](images/Ch11_2.png)

# Chapter 11-3
![images/Ch11_3.png](images/Ch11_3.png)

# Chapter 12
![images/Ch12.png](images/Ch12.png)

# Chapter 13
![images/Ch13.png](images/Ch13.png)

# Chapter 14
![images/Ch14.png](images/Ch14.png)

# Chapter 9 - transparent, different RIs
![images/3GlassBalls.png](images/3GlassBalls.png)

# Scene base on Peter Shirley's work
![images/Glass.png](images/glass.png)
