import codecs
import xlsxwriter
#import pandas as pd
import os
import uuid
import numpy as np

import math

def rad2grad(rad):
    return rad*200.0/math.pi
def grad2rad(grad):
    return grad*math.pi/200.0

def deftero(t1,o1):
    x1=t1[0]
    y1=t1[1]
    x2=o1[0]
    y2=o1[1]
    dx=x2-x1
    dy=y2-y1
    D=math.sqrt((dx**2)+(dy**2))
    if (D==0):
        return False
    if (dy==0):
        if (dx>0):
            return [D,100]
        if (dx>0):
            return [D,300]
    a=rad2grad(math.atan(abs(dx/dy)))
    if (dy>0):
        if (dx>0):
            return [D,a]
        if (dx==0):
            return [D,0]
        if (dx<0):
            return [D,400-a]
    if (dy<0):
        if (dx>0):
            return [D,200-a]
        if (dx==0):
            return [D,200]
        if (dx<0):
            return [D,200+a]
import codecs       

def rotate_point(P,theta):
    g=deftero([0,0],P)
    angle=g[1]+theta
    rad=grad2rad(angle)
    X=g[0]*math.sin(rad)
    Y=g[0]*math.cos(rad)
    return [X,Y]

p=rotate_point([100,100],1)
#print(p)
import copy

def area(pointlist):
    npl=copy.deepcopy(pointlist)
    npl=[]
    for i in pointlist:
        npl.append(i)
    e=0
    npl.append(pointlist[0])
    n=len(npl)
    for i in range(len(npl)-1):
        xi=npl[i][0]
        xi1=npl[i+1][0]
        yi=npl[i][1]
        yi1=npl[i+1][1]
        s=(xi+xi1)*(yi1-yi)
        e=e+s
        #print(i,e)
    return abs(e/2)
        
def eq(A,B):
    x1=A[0]
    x2=B[0]
    y1=A[1]
    y2=B[1]
    return [y1-y2,x2-x1,x1*y2-y1*x2]
 
def ver(e,point):
    return [-1/e[0],-1,point[1]+point[0]/e[0]]

def best_fit(list_of_points):
    global sx
    global sy
    global sxy
    global sx2
    global a
    global b
    global n
    n=len(list_of_points)
    all=np.array(list_of_points)
    sx=all[:,0].sum()
    sy=all[:,1].sum()
    s2=all[:,0]*all[:,0]
    sx2=s2.sum()
    #print(sx2)
    sxy=(all[:,0]*all[:,1]).sum()
    a=(n*sxy-sx*sy)/(n*sx2-sx*sx)
    b=(sx2*sy-sx*sxy)/(n*sx2-sx*sx)
    return a,-1,b,

def add_distance(line,distance,beta_gt_gamma):
    a=line[0]
    b=line[1]
    c=line[2]
    if beta_gt_gamma==False:
        c=c-distance*math.sqrt(1+a*a)
    else:
        c=c+distance*math.sqrt(1+a*a)
    return [a,b,c]

def line_intersection2(line1, line2):
    a1=line1[0]
    b1=line1[1]
    c1=line1[2]
    a2=line2[0]
    b2=line2[1]
    c2=line2[2]
    x=(b1*c2-b2*c1)/(a1*b2-a2*b1)
    y=(c1*a2-c2*a1)/(a1*b2-a2*b1)
    return x, y

def main(DX,DY,dy,theta,exam_id,folder):
    A=[-8.771,18.293]
    B=[25.046,18.832]
    C=[-25.5,38.197]
    D=[8.5,32.914]
    E=[10.423,32.922]
    F=[10.423,32.061]
    G=[20.5,32.149]
    H=[20.5,40.704]
    I=[8.5,40.409]
    J=[-12.112,40.573]
    K=[-12.113,32.216]
    L=[-20.493,32.017]
    M=[-20.442,40.376]
    N=[23.5,-17.142]
    O=[-23.5,-17.142]
    P=[8.5,-2.858]
    Q=[23.5,-2.858]
    R=[23.5,17.142]
    S=[8.5,17.142]
    T=[8.5,4.142]
    U=[-2.5,4.142]
    V=[-2.5,17.142]
    W=[-23.5,2.142]
    X=[-11.5,2.142]
    Y=[-11.5,4.142]
    Z=[-11.5,17.142]
    A1=[-23.5,17.142]
    A2=[-23.5,38.197]
    A3=[-23.5,29.142]
    A4=[23.5,41.564]
    A5=[23.5,29.142]
    A6=[25.5,28.142]
    A7=[24.5,27.142]
    A8=[-24.5,27.142]
    A9=[-25.5,28.142]
    A10=[25.5,18.142]
    A11=[24.5,19.142]
    A12=[-24.5,19.142]
    A13=[-25.5,18.142]
    A14=[-25.5,-18.142]
    A15=[-24.5,-19.142]
    A16=[24.5,-19.142]
    A17=[25.5,-18.142]
    pointslist=[A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,A1,A2,A3,A4,A5,A6,A7,A8,A9,A10,A11,A12,A13,A14,A15,A16,A17]
    rotated=[]
    workbook = xlsxwriter.Workbook('c:/autolisp/examsJan24/{0}/exams{1}.xlsx'.format(folder,exam_id))
    worksheet = workbook.add_worksheet()

   
    file5 = codecs.open("c:/autolisp/examsJan24/exams{1}.txt".format(folder,exam_id),"w", "utf-8")
    file =open("out.txt","w")
    workbook = xlsxwriter.Workbook('c:/autolisp/examsJan24/{0}/exams{1}.xlsx'.format(folder,exam_id))
    worksheet = workbook.add_worksheet()
    worksheet.write(2, 1,"Πίνακας συντεταγμένων και μετρήσεων ")
    worksheet.write(3, 0,     "κωδικοί")
    worksheet.write(3, 1,     "Συντεταγμενες")
    worksheet.write(4, 0,     "id")
    worksheet.write(4, 1,     "X")
    worksheet.write(4, 2,     "Y")
    T[1]=T[1]-dy
    U[1]=U[1]-dy
    Y[1]=Y[1]-dy
    for idx,P in enumerate(pointslist):
        np=rotate_point(P,theta)
        rotated.append(np)
        if idx<26:
            file.write("{0} {1:.3f} {2:.3f}\n".format(chr(idx+65),np[0]+DX,np[1]+DY))
            worksheet.write(idx+4, 0, "{0}".format(chr(idx+65)))
        else:
            file.write("A{0} {1:.3f} {2:.3f}\n".format(idx-25,np[0]+DX,np[1]+DY))
            worksheet.write(idx+4, 0, "A{0}".format(idx-25))
        worksheet.write(idx+4, 1, round(np[0]+DX,3))
        worksheet.write(idx+4, 2, round(np[1]+DY,3))
    file.close()
    L1=[round(rotate_point(L,theta)[0],3)+DX,round(rotate_point(L,theta)[1],3)+DY]
    K1=[round(rotate_point(K,theta)[0],3)+DX,round(rotate_point(K,theta)[1],3)+DY]
    F1=[round(rotate_point(F,theta)[0],3)+DX,round(rotate_point(F,theta)[1],3)+DY]
    G1=[round(rotate_point(G,theta)[0],3)+DX,round(rotate_point(G,theta)[1],3)+DY]
    L1=[round(rotate_point(L,theta)[0],3)+DX,round(rotate_point(L,theta)[1],3)+DY]
    K1=[round(rotate_point(K,theta)[0],3)+DX,round(rotate_point(K,theta)[1],3)+DY]
    F1=[round(rotate_point(F,theta)[0],3)+DX,round(rotate_point(F,theta)[1],3)+DY]
    G1=[round(rotate_point(G,theta)[0],3)+DX,round(rotate_point(G,theta)[1],3)+DY]
    Z1=[round(rotate_point(Z,theta)[0],3)+DX,round(rotate_point(Z,theta)[1],3)+DY]
    V1=[round(rotate_point(V,theta)[0],3)+DX,round(rotate_point(V,theta)[1],3)+DY]
    U1=[round(rotate_point(U,theta)[0],3)+DX,round(rotate_point(U,theta)[1],3)+DY]
    Y1=[round(rotate_point(Y,theta)[0],3)+DX,round(rotate_point(Y,theta)[1],3)+DY]
    Z1=[round(rotate_point(Z,theta)[0],3)+DX,round(rotate_point(Z,theta)[1],3)+DY]
    V1=[round(rotate_point(V,theta)[0],3)+DX,round(rotate_point(V,theta)[1],3)+DY]
    U1=[round(rotate_point(U,theta)[0],3)+DX,round(rotate_point(U,theta)[1],3)+DY]
    Y1=[round(rotate_point(Y,theta)[0],3)+DX,round(rotate_point(Y,theta)[1],3)+DY]
    A1=[round(rotate_point(A,theta)[0],3)+DX,round(rotate_point(A,theta)[1],3)+DY]
    B1=[round(rotate_point(B,theta)[0],3)+DX,round(rotate_point(B,theta)[1],3)+DY]
    bb1=best_fit([L1,K1,F1,G1])
    bb2=add_distance(bb1,18,False)
    #print(bb1)
    #print(bb2)
    line1=eq(Z1,Y1)
    line2=eq(U1,V1)
    TT1=line_intersection2(bb2,line1)
    TT2=line_intersection2(bb2,line2)
    #print(TT1)
    #print(TT2)
    #print(A1)
    #print(B1)
    Em=area([Z1,V1,U1,Y1])
    g0=deftero(A1,B1)
    g1=deftero(A1,TT1)
    g2=deftero(A1,TT2)
    a1=g1[1]-g0[1]
    a2=g2[1]-g0[1]
    #print(g0)
    #print(g1)
    #print(g2)
    file5.write("Aσκηση\n")
    file5.write('Κατεβάστε το αρχείο των μετρήσεων από τον <a href="http://155.207.25.31/exams/{0}/exams{1}.xlsx">δεσμό</a> '.format(folder,exam_id))
    file5.write('για την απόδοση του σχεδίου σύμφωνα με το <a href="http://155.207.25.31/exams/kroki.pdf">κροκί</a>\n'.format(folder)) 
    file5.write('\nΥπολογίστε την εξίσωση της ευθείας της οικοδομικής γραμμής στο όμορο ΟΤ στη μορφή y=a1x+b1 και δώστε παρακάτω τις παραμέτρους:\na1 {') 

    file5.write("1:NUMERICAL:={0:.7f}".format(bb1[0]))
    file5.write(":0.000002}")
    file5.write("\nb1 {")
    file5.write("2:NUMERICAL:={0:.7f}".format(bb1[2]))
    file5.write(":0.02}\n")
    file5.write('\nΌπως επίσης και της ευθείας της οικοδομικής γραμμής στη βόρεια πλευρά του ΟΤ ενδιαφέροντος στη μορφή y=a2x+b2 και δώστε παρακάτω τις παραμέτρους:\na2 {') 
    file5.write("3:NUMERICAL:={0:.7f}".format(bb2[0]))
    file5.write(":0.000002}")
    file5.write("\nb2 {")
    file5.write("4:NUMERICAL:={0:.7f}".format(bb2[2]))
    file5.write(":0.02}\n")
    file5.write('\nΕπιπλέον υπολογίστε τις συντεταγμένες των σημείων τομής της οικοδομικής γραμμής με τα όρια του σκιαγραμμισμένου οικοπέδου Τ1 και Τ2:\nΤ1X {') 
    file5.write("5:NUMERICAL:={0:.7f}".format(TT1[0]))
    file5.write(":0.02}\n")
    file5.write("\nΤ1Y {")
    file5.write("6:NUMERICAL:={0:.7f}".format(TT1[1]))
    file5.write(":0.02}\n\nT2X {")
    file5.write("7:NUMERICAL:={0:.7f}".format(TT2[0]))
    file5.write(":0.02}\n")
    file5.write("\nΤ2Y {")
    file5.write("8:NUMERICAL:={0:.7f}".format(TT2[1]))
    file5.write(":0.02}\n")
    file5.write('\nYπολογίστε το εμβαδόν του σκιαγραμμισμένου οικοπέδου:\nΕ {') 
    file5.write("9:NUMERICAL:={0:.7f}".format(Em))
    file5.write(":0.02}\n")
    file5.write('\nΤέλος υπολογίστε τις συντεταγμένες των πολικών συντεταγμένων για τη χάρξη των σημείων Τ1 και Τ2 με στάση οργάνου το Α και προσανατολισμό το σημείο Β:\nΤ1D {') 
    file5.write("10:NUMERICAL:={0:.7f}".format(g1[0]))
    file5.write(":0.02}\n")
    file5.write("\nΤ1θ {")
    file5.write("11:NUMERICAL:={0:.7f}".format(a1))
    file5.write(":0.02}\n")
    file5.write("\nΤ2D {")
    file5.write("13:NUMERICAL:={0:.7f}".format(g2[0]))
    file5.write(":0.02}\n")   
    file5.write("\nΤ2θ {")
    file5.write("14:NUMERICAL:={0:.7f}".format(a2))
    file5.write(":0.02}\n")  
    file5.close()
    workbook.close()


def main2(DX,DY,dy,theta,exam_id,folder):
    A=[-8.771,18.293]
    B=[25.046,18.832]
    C=[-25.5,38.197]
    D=[8.5,32.914]
    E=[10.423,32.922]
    F=[10.423,32.061]
    G=[20.5,32.149]
    H=[20.5,40.704]
    I=[8.5,40.409]
    J=[-12.112,40.573]
    K=[-12.113,32.216]
    L=[-20.493,32.017]
    M=[-20.442,40.376]
    N=[23.5,-17.142]
    O=[-23.5,-17.142]
    P=[8.5,-2.858]
    Q=[23.5,-2.858]
    R=[23.5,17.142]
    S=[8.5,17.142]
    T=[8.5,4.142]
    U=[-2.5,4.142]
    V=[-2.5,17.142]
    W=[-23.5,2.142]
    X=[-11.5,2.142]
    Y=[-11.5,4.142]
    Z=[-11.5,17.142]
    A1=[-23.5,17.142]
    A2=[-23.5,38.197]
    A3=[-23.5,29.142]
    A4=[23.5,41.564]
    A5=[23.5,29.142]
    A6=[25.5,28.142]
    A7=[24.5,27.142]
    A8=[-24.5,27.142]
    A9=[-25.5,28.142]
    A10=[25.5,18.142]
    A11=[24.5,19.142]
    A12=[-24.5,19.142]
    A13=[-25.5,18.142]
    A14=[-25.5,-18.142]
    A15=[-24.5,-19.142]
    A16=[24.5,-19.142]
    A17=[25.5,-18.142]
    pointslist=[A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,A1,A2,A3,A4,A5,A6,A7,A8,A9,A10,A11,A12,A13,A14,A15,A16,A17]
    rotated=[]
    workbook = xlsxwriter.Workbook('c:/autolisp/examsJan24/{0}/exams{1}.xlsx'.format(folder,exam_id))
    worksheet = workbook.add_worksheet()

   
    file5 = codecs.open("c:/autolisp/examsJan24/exams{1}.txt".format(folder,exam_id),"w", "utf-8")
    file =open("out.txt","w")
    workbook = xlsxwriter.Workbook('c:/autolisp/examsJan24/{0}/exams{1}.xlsx'.format(folder,exam_id))
    worksheet = workbook.add_worksheet()
    worksheet.write(2, 1,"Πίνακας συντεταγμένων και μετρήσεων ")
    worksheet.write(3, 0,     "κωδικοί")
    worksheet.write(3, 1,     "Συντεταγμενες")
    worksheet.write(4, 0,     "id")
    worksheet.write(4, 1,     "X")
    worksheet.write(4, 2,     "Y")
    T[1]=T[1]-dy
    U[1]=U[1]-dy
    Y[1]=Y[1]-dy
    for idx,P in enumerate(pointslist):
        np=rotate_point(P,theta)
        rotated.append(np)
        if idx<26:
            file.write("{0} {1:.3f} {2:.3f}\n".format(chr(idx+65),np[0]+DX,np[1]+DY))
            worksheet.write(idx+4, 0, "{0}".format(chr(idx+65)))
        else:
            file.write("A{0} {1:.3f} {2:.3f}\n".format(idx-25,np[0]+DX,np[1]+DY))
            worksheet.write(idx+4, 0, "A{0}".format(idx-25))
        worksheet.write(idx+4, 1, round(np[0]+DX,3))
        worksheet.write(idx+4, 2, round(np[1]+DY,3))
    file.close()
    L1=[round(rotate_point(L,theta)[0],3)+DX,round(rotate_point(L,theta)[1],3)+DY]
    K1=[round(rotate_point(K,theta)[0],3)+DX,round(rotate_point(K,theta)[1],3)+DY]
    F1=[round(rotate_point(F,theta)[0],3)+DX,round(rotate_point(F,theta)[1],3)+DY]
    G1=[round(rotate_point(G,theta)[0],3)+DX,round(rotate_point(G,theta)[1],3)+DY]
    L1=[round(rotate_point(L,theta)[0],3)+DX,round(rotate_point(L,theta)[1],3)+DY]
    K1=[round(rotate_point(K,theta)[0],3)+DX,round(rotate_point(K,theta)[1],3)+DY]
    F1=[round(rotate_point(F,theta)[0],3)+DX,round(rotate_point(F,theta)[1],3)+DY]
    G1=[round(rotate_point(G,theta)[0],3)+DX,round(rotate_point(G,theta)[1],3)+DY]
    Z1=[round(rotate_point(Z,theta)[0],3)+DX,round(rotate_point(Z,theta)[1],3)+DY]
    V1=[round(rotate_point(V,theta)[0],3)+DX,round(rotate_point(V,theta)[1],3)+DY]
    U1=[round(rotate_point(U,theta)[0],3)+DX,round(rotate_point(U,theta)[1],3)+DY]
    Y1=[round(rotate_point(Y,theta)[0],3)+DX,round(rotate_point(Y,theta)[1],3)+DY]
    Z1=[round(rotate_point(Z,theta)[0],3)+DX,round(rotate_point(Z,theta)[1],3)+DY]
    V1=[round(rotate_point(V,theta)[0],3)+DX,round(rotate_point(V,theta)[1],3)+DY]
    U1=[round(rotate_point(U,theta)[0],3)+DX,round(rotate_point(U,theta)[1],3)+DY]
    Y1=[round(rotate_point(Y,theta)[0],3)+DX,round(rotate_point(Y,theta)[1],3)+DY]
    A1=[round(rotate_point(A,theta)[0],3)+DX,round(rotate_point(A,theta)[1],3)+DY]
    B1=[round(rotate_point(B,theta)[0],3)+DX,round(rotate_point(B,theta)[1],3)+DY]
    bb1=best_fit([L1,K1,F1,G1])
    bb2=add_distance(bb1,18,False)
    #print(bb1)
    #print(bb2)
    line1=eq(Z1,Y1)
    line2=eq(U1,V1)
    TT1=line_intersection2(bb2,line1)
    TT2=line_intersection2(bb2,line2)
    #print(TT1)
    #print(TT2)
    #print(A1)
    #print(B1)
    Em=area([Z1,V1,U1,Y1])
    g0=deftero(A1,B1)
    g1=deftero(A1,TT1)
    g2=deftero(A1,TT2)
    a1=g1[1]-g0[1]
    a2=g2[1]-g0[1]
    #print(g0)
    #print(g1)
    #print(g2)
    file5.write("Aσκηση\n")
    file5.write('Κατεβάστε το αρχείο των μετρήσεων από τον <a href="http://155.207.25.85:8080/examsJan24/{0}/exams{1}.xlsx">δεσμό</a> '.format(folder,exam_id))
    file5.write('για την απόδοση του σχεδίου σύμφωνα με το <a href="http://155.207.25.85:8080/examsJan24/kroki.pdf">κροκί</a>\n'.format(folder)) 
    file5.write('\nΥπολογίστε την εξίσωση της ευθείας της οικοδομικής γραμμής στο όμορο ΟΤ στη μορφή y=a1x+b1 και δώστε παρακάτω τις παραμέτρους:\na1 {') 

    file5.write("1:NUMERICAL:={0:.7f}".format(bb1[0]))
    file5.write(":0.000002}")
    file5.write("\nb1 {")
    file5.write("2:NUMERICAL:={0:.7f}".format(bb1[2]))
    file5.write(":0.02}\n")
    file5.write('\nΌπως επίσης και της ευθείας της οικοδομικής γραμμής στη βόρεια πλευρά του ΟΤ ενδιαφέροντος στη μορφή y=a2x+b2 και δώστε παρακάτω τις παραμέτρους:\na2 {') 
    file5.write("3:NUMERICAL:={0:.7f}".format(bb2[0]))
    file5.write(":0.000002}")
    file5.write("\nb2 {")
    file5.write("4:NUMERICAL:={0:.7f}".format(bb2[2]))
    file5.write(":0.02}\n")
    file5.write('\nΕπιπλέον υπολογίστε τις συντεταγμένες των σημείων τομής της οικοδομικής γραμμής με τα όρια του σκιαγραμμισμένου οικοπέδου Τ1 και Τ2:\nΤ1X {') 
    file5.write("5:NUMERICAL:={0:.7f}".format(TT1[0]))
    file5.write(":0.02}\n")
    file5.write("\nΤ1Y {")
    file5.write("6:NUMERICAL:={0:.7f}".format(TT1[1]))
    file5.write(":0.02}\n\nT2X {")
    file5.write("7:NUMERICAL:={0:.7f}".format(TT2[0]))
    file5.write(":0.02}\n")
    file5.write("\nΤ2Y {")
    file5.write("8:NUMERICAL:={0:.7f}".format(TT2[1]))
    file5.write(":0.02}\n")
    file5.write('\nYπολογίστε το εμβαδόν του σκιαγραμμισμένου οικοπέδου:\nΕ {') 
    file5.write("9:NUMERICAL:={0:.7f}".format(Em))
    file5.write(":0.02}\n")
    file5.write('\nΤέλος υπολογίστε τις συντεταγμένες των πολικών συντεταγμένων για τη χάρξη των σημείων Τ1 και Τ2 με στάση οργάνου το Α και προσανατολισμό το σημείο Β:\nΤ1D {') 
    file5.write("10:NUMERICAL:={0:.7f}".format(g1[0]))
    file5.write(":0.02}\n")
    file5.write("\nΤ1θ {")
    file5.write("11:NUMERICAL:={0:.7f}".format(a1))
    file5.write(":0.02}\n")
    file5.write("\nΤ2D {")
    file5.write("13:NUMERICAL:={0:.7f}".format(g2[0]))
    file5.write(":0.02}\n")   
    file5.write("\nΤ2θ {")
    file5.write("14:NUMERICAL:={0:.7f}".format(a2))
    file5.write(":0.02}\n")  
    file5.close()
    workbook.close()

import random
from numpy.random import randint
for i in range(1,53):
    folder=str(uuid.uuid4())
    os.mkdir("c:/autolisp/examsJan24/"+folder)
    DX=randint(100,200)
    DY=randint(100,200)
    #print(DX,DY)
    #main(DX,DY,random.uniform(0,2),random.uniform(0,5),i,folder) #για τον Clara
    main2(DX,DY,random.uniform(0,2),random.uniform(0,5),i,folder) #για τον Synology
