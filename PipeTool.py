# SOURCE:https://www.alibre.com/forum/index.php?threads/ai-scripting-new-tools-into-alibre.26141/#post-179976
# Alibre Script: Advanced Sweep (Single Sketch Method)
# Creates a solid or hollow sweep from a single profile sketch.
import sys

# Establish handles for the main Alibre objects
MyPart = CurrentPart()
Win = Windows()

# --- 1. Define and Display the Menu ---
Units.Current = UnitTypes.Millimeters

# Define the layout of the static dialog window
Options = []
Options.append(['Path Sketch', WindowsInputTypes.Sketch, None])
Options.append(['Profile Type', WindowsInputTypes.StringList, ['Circle', 'Square']])
Options.append(['Size (mm)', WindowsInputTypes.Real, 10.0])
Options.append(['Hollow', WindowsInputTypes.Boolean, False])
Options.append(['Thickness (mm)', WindowsInputTypes.Real, 1.0])

# Show the options dialog
Values = Win.OptionsDialog('Create Advanced Sweep', Options, 220)

# Exit the script if the user cancels the dialog
if Values is None or Values[0] is None:
  print "Operation cancelled or no path sketch selected."
  sys.exit()

# --- 2. Process Inputs ---
PathSketch = Values[0]
ProfileTypeIndex = Values[1]
ProfileSize = Values[2]
IsHollow = Values[3]
Thickness = Values[4]

# --- Define the Profile Plane (Works for Lines and Splines) ---
FirstFigure = PathSketch.Figures[0]
Is3DSketch = '3D' in PathSketch.GetType().Name

if hasattr(FirstFigure, 'GetPointAt'): # Spline logic
  if Is3DSketch:
    StartPoint3D = FirstFigure.GetPointAt(0.0)
    NextPoint3D = FirstFigure.GetPointAt(0.001)
  else:
    StartPoint2D = FirstFigure.GetPointAt(0.0)
    NextPoint2D = FirstFigure.GetPointAt(0.001)
    StartPoint3D = PathSketch.PointtoGlobal(StartPoint2D[0], StartPoint2D[1])
    NextPoint3D = PathSketch.PointtoGlobal(NextPoint2D[0], NextPoint2D[1])
else: # Line logic
  if Is3DSketch:
    StartPoint3D = FirstFigure.StartPoint
    NextPoint3D = FirstFigure.EndPoint
  else:
    StartPoint2D = FirstFigure.StartPoint
    NextPoint2D = FirstFigure.EndPoint
    StartPoint3D = PathSketch.PointtoGlobal(StartPoint2D[0], StartPoint2D[1])
    NextPoint3D = PathSketch.PointtoGlobal(NextPoint2D[0], NextPoint2D[1])

DirectionVector = [NextPoint3D[i] - StartPoint3D[i] for i in range(3)]
ProfilePlane = MyPart.AddPlane('SweepProfilePlane', DirectionVector, StartPoint3D)

# --- 3. Create a Single Profile Sketch ---
ProfileSketch = MyPart.AddSketch('SweepProfile', ProfilePlane)

# Draw the outer profile shape based on the index from the dropdown
if ProfileTypeIndex == 0: # Circle
  OuterSize = ProfileSize
  ProfileSketch.AddCircle(0, 0, OuterSize, False)
elif ProfileTypeIndex == 1: # Square
  OuterHalf = ProfileSize / 2.0
  ProfileSketch.AddRectangle(-OuterHalf, -OuterHalf, OuterHalf, OuterHalf, False)

# If hollow is checked, draw the inner profile ON THE SAME SKETCH
if IsHollow and Thickness > 0 and Thickness < (ProfileSize / 2.0):
  if ProfileTypeIndex == 0: # Circle
    InnerSize = ProfileSize - (Thickness * 2)
    ProfileSketch.AddCircle(0, 0, InnerSize, False)
  elif ProfileTypeIndex == 1: # Square
    InnerSize = ProfileSize - (Thickness * 2)
    InnerHalf = InnerSize / 2.0
    ProfileSketch.AddRectangle(-InnerHalf, -InnerHalf, InnerHalf, InnerHalf, False)

# --- 4. Perform the Sweep ---
# Regenerate the part to ensure the sketch (with one or two loops) is finalized
MyPart.Regenerate()

# Create the sweep feature using the single profile sketch
MyPart.AddSweepBoss('Sweep', ProfileSketch, PathSketch, False, MyPart.EndCondition.EntirePath, None, 0, 0, False)

print "Script finished."