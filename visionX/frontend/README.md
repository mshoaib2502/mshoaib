# BACKEND
pip install fastapi uvicorn opencv-python python-multipart
uvicorn main:app --reload --port 8000

# FRONTEND
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
npm create vite@latest frontend -- --template react
cd frontend
npm i konva react-konva
npm run dev

# Logic from github copilot

The provided code is a React component named [`Annotator`], which is designed for annotating frames from a video. It utilizes the [`react-konva`] library to render a canvas where users can draw bounding boxes around objects in the video frames. The component accepts two props: [`videoId`], which identifies the video being annotated, and [`frames`], an array of frame names corresponding to the video.

At the beginning of the component, several state variables are defined using the [`useState`] hook. [`idx`] keeps track of the current frame index, [`boxes`] stores the annotations (bounding boxes) drawn by the user, and [`drawing`] holds the current drawing state, including the position and dimensions of the box being drawn. The [`frameName`] variable is derived from the [`frames`] array based on the current index.

The [`useEffect`] hook is employed to fetch annotations from a backend API whenever the [`frameName`] changes. This API call retrieves existing annotations for the current frame and updates the [`boxes`] state with the fetched data. The [`save`] function allows users to save their annotations back to the server by sending a POST request with the current boxes in JSON format.

Mouse event handlers are defined to manage the drawing of bounding boxes. [`handleMouseDown`] initializes the drawing state when the user clicks on the canvas, capturing the starting coordinates. [`handleMouseMove`] updates the dimensions of the box as the mouse moves, while [`handleMouseUp`] finalizes the box if its dimensions are significant (greater than 10 pixels) and adds it to the [`boxes`] array. If the drawing is completed, the [`drawing`] state is reset.

The component's return statement renders a user interface that includes navigation buttons for moving between frames, a display of the current frame index, and a button to save annotations. The [`Stage`] component from [`react-konva`] is used to create the drawing area, where the background image is set to the current frame. The [`Layer`] component contains the drawn boxes and their labels, with a visual distinction for the currently drawn box using a dashed red stroke. Overall, this component provides an interactive way for users to annotate video frames effectively.