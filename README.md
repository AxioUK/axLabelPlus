# axLabelPlus v1.6.12

Fork of Awesome LabelPlus VB6 UserControl from Leandro Ascierto

## Properties

:star:: additions versus original LabelPlus​

| Properties                      |   Value   | Description                                                  |
| ------------------------------- | :-------: | ------------------------------------------------------------ |
| AutoSize                        |  Boolean  | *allows the control to manage its measurements according to the content string* |
| BackAcrylicBlur                 |  Boolean  | *gives the control a transparent blurred glass effect*       |
| BackColor                       | OLE_Color | *control background color*                                   |
| BackColorOpacity                |  Integer  | *control background color opacity*                           |
| :star: BackColorPress        | OLE_Color | *control background color on click*                          |
| :star: BackColorPressOpacity |  Integer  | *control background color opacity on click*                  |
| Border                          |  Boolean  | *control border*                                             |
| BorderColor                     | OLE_Color | *control border color*                                       |
| BorderColorOpacity              |  Integer  | *control border color opacity*                               |
| BorderCornerBottomLeft          |  Integer  | *radius for bottom-left round corner*                        |
| BorderCornerBottomRight         |  Integer  | *radius for bottom-right round corner*                       |
| BorderCornerLeftTop             |  Integer  | *radius for top-left round corner*                           |
| BorderCornerRightTop            |  Integer  | *radius for top-right  round corner*                         |
| BorderPosition                  | Constant  | *indicates if edge line is drawn inside, outside or center of the border* |
| BorderWidth                     |  Integer  | *width of the edge line*                                     |
| CallOut                         |  Boolean  | *enable the callout triangle (to simulate a bubble text or Tooltip text)* |
| CallOutAlign                    | Constant  | *position of the callout triangle on the corner especified*  |
| CallOutCustomPosition           |  Integer  | *value to positioning callout triangle when  "coCustomPosition" is indicated* |
| CallOutLen                      |  Integer  | *length of the callout triangle*                             |
| CallOutPosicion                 | Constant  | *position of the callout triangle on the side especified*    |
| CallOutRightTriangle            |  Boolean  | *indicate if the callout is a rect-triangle or not*          |
| CallOutWidth                    |  Integer  | *width of the callout triangle border line*                  |
| Caption1                        |  String   | *text string of the first caption*                           |
| Caption1Font                    |  StdFont  | *font settings to the first caption*                         |
| Caption1Forecolor               | OLE_color | *font color to the first caption*                            |
| Caption1ForeColorOpacity        |  Integer  | *font color opacity to the first caption*                    |
| Caption1PaddingX                |  Integer  | *horizontal scrolling of the first caption*                  |
| Caption1PaddingY                |  Integer  | *vertical scrolling of the first caption*                    |
| Caption1WordWrap                |  Boolean  | *enables wordwrap to first caption if text string is too long* |
| :star: Caption2              |  String   | *text string of the second caption*                          |
| :star: Caption2Font          |  StdFont  | *font settings to the second caption*                        |
| :star: Caption2Forecolor     | OLE_color | *font color to the second caption*                           |
| :star: Caption2ForeColorOpacity |  Integer  | *font color opacity to the second  caption*                  |
| :star: Caption2PaddingX     |  Integer  | *horizontal scrolling of the second caption*                 |
| :star: Caption2PaddingY     |  Integer  | *vertical scrolling of the second caption*                   |
| :star: Caption2WordWrap     |  Boolean  | *enables wordwrap to second caption if text string is too long* |
| CaptionAlignmentH               | Constant  | *horizontal alignment of the both captions (it does not superseed the assigned PaddingX values, so it checks its values to align properly)* |
| CaptionAlignmentV               | Constant  | *vertical alignment of the both captions (it does not superseed the assigned PaddingY values, so it checks its values to align properly)* |
| CaptionAngle                    |  Integer  | *inclination angle for the captions*                         |
| CaptionBorderColor              | OLE_Color | *border color of text string in captions*                    |
| CaptionBorderWidth              |  Integer  | *border width of text string in captions*                    |
| CaptionShadow                   |  Boolean  | *shadow of text strings in captions*                         |
| CaptionShowPrefix               |  Boolean  | *indicate how process the ampersand (&) character when it is indicated in captions (ej: &Function... show <u>F</u>unction..)* |
| CaptionTriming                  | Constant  | *indicates how to process long text strings in captions*     |
| :star: ChangeColorOnClick   |  Boolean  | *enable change of color of the control when click on it (BackColorPress, GradientColorP1, GradientColorP2)* |
| :star: ChangeOnMouseOver    | Constant  | *indicate styles of color change when mouse passed on control* |
| :star: ColorOnMouseOver      | OLE_Color | *indicate color to change if ChangeOnMouseOver is setted*    |
| :star:ColorOpacityOnMouseOver |  Integer  | *color opacity to* ColorOnMouseOver                          |
| :star: ​CrossPosition       | Constant  | *indicate the position to place Cross to "close" the control* |
| :star: CrossVisible          |  Boolean  | *enable the cross "to close" (hide) the control (it change the state of VISIBLE property)* |
| :star: ForeColorOnPress      | OLE_Color | *color of the font (in both captions) when control is clicked and ChangeColorOnClick is setted to     eChangeCaption1, eChangeCaption2, eChangeCaptions, eChangeCaptionIcon, eChangeCaptionBorder or eChangeCaptionHotLine* |
| :star:ForeColorOnPressOpacity |  Integer  | *color opacity to* ForeColorOnPress                          |
| :star: GlowColor             | OLE_Color | *color border for Glowing effect*                            |
| :star: Glowing              |  Boolean  | *enable/start Glowing effect*                                |
| :star: GlowSpeed            |  Integer  | *speed of the pulse in glowing effect*                       |
| :star: GlowTiks              |  Integer  | *indicates how many pulses to perform before stopping the effect* |
| Gradient                        |  Boolean  | *enable gradient of two colors on the control backgound*     |
| GradientAngle                   |  Integer  | *inclination angle of the gradient*                          |
| GradientColor1                  | OLE_Color | *first color of the gradient*                                |
| GradientColor1Opacity           |  Integer  | *opacity of the first color*                                 |
| GradientColor2                  | OLE_Color | *last color of the gradient*                                 |
| GradientColor2Opacity           |  Integer  | *opacity of the last color*                                  |
| :star: GradientColorP1      | OLE_Color | *first color of the gradient when control is clicked*        |
| :star: GradientColorP1Opacity |  Integer  | *opacity of the first color of the gradient when control is clicked* |
| :star: GradientColorP2       | OLE_Color | *last color of the gradient when control is clicked*         |
| :star: ​GradientColorP2Opacity |  Integer  | *opacity of the last color of the gradient when control is clicked* |
| HotLine                         |  Boolean  | *enable a color bar on one side of the control*              |
| HotLineColor                    | OLE_Color | *color of the HotLine bar*                                   |
| HotLineColorOpacity             |  Integer  | *color opacity of the HotLine bar*                           |
| HotLinePosition                 | Constant  | *position of the HotLine color bar*                          |
| HotLineWidth                    |  Integer  | *width of the color bar*                                     |
| IconAlignmentH                  | Constant  | *indicate the horizontal alignment for iconchar*             |
| IconAlignmentV                  | Constant  | *indicate the vertical alignment for iconchar*               |
| IconCharCode                    |   Long    | *the icon char code that we want to show (according to the icon font to use)* |
| IconFont                        |  StdFont  | *set the icon font (recomended [IcoFont.ttf](https://icofont.com/icons))* |
| IconForeColor                   | OLE_Color | *set color of the icon char*                                 |
| IconOpacity                     |  Integer  | *set the opacity color value of the icon char*               |
| IconPaddingX                    |  Integer  | *horizontal scrolling of the icon char*                      |
| IconPaddingY                    |  Integer  | *vertical scrolling of the icon char*                        |
| Image                           | IPicture  | *get the picture contained in the control*                   |
| IsMouseEnter                    |  Boolean  | *return* TRUE *when the mouse enter to the control area, return* FALSE *on exit* |
| IsMouseOver                     |  Boolean  | *return* TRUE *if mouse is over control area, return* FALSE *if not* |
| :star: OptionBehavior       |  Boolean  | *allows the control to function as an OptionButton, returning* TRUE *when clicked and passing the other controls that have OptionBehavior enabled to* FALSE |
| PictureAlignmentH               | Constant  | *indicate the horizontal alignment of the picture*           |
| PictureAlignmentV               | Constant  | *indicate the vertical alignment of the picture*             |
| PictureAngle                    |  Integer  | *set the inclination angle of the picture*                   |
| PictureBrightness               |  Integer  | *set the brightness of the picture*                          |
| PictureColor                    | OLE_Color | *set a single color to the picture when* PictureColorize *is enabled* |
| PictureColorize                 |  Boolean  | *enable the effect of putting the image in one color*        |
| PictureContrast                 |  Integer  | *set the contrast of the picture*                            |
| PictureExist                    |  Boolean  | *return* TRUE *if exist a picture in the control*            |
| PictureGetHeight                |   Long    | *return the height of the picture setted in the control*     |
| PictureGetWidth                 |   Long    | *return the width of the picture setted in the control*      |
| PictureGrayScale                |  Boolean  | *enable the effect of grayscaling the picture*               |
| PictureOpacity                  |  Integer  | *set the opacity of the picture*                             |
| PicturePaddingX                 |  Integer  | *horizontal scrolling of the picture*                        |
| PicturePaddingY                 |  Integer  | *vertical scrolling of the picture*                          |
| PictureSetHeight                |   Long    | *set/override the height of the picture*                     |
| PictureSetWidth                 |   Long    | *set/override the width of the picture*                      |
| PictureShadow                   |  Boolean  | *enable the picture shadow*                                  |
| Shadow                          |  Boolean  | *enable control shadow*                                      |
| ShadowColor                     | OLE_Color | *set the color of the control shadow*                        |
| ShadowColorOpacity              |  Integer  | *set the color opacity value*                                |
| ShadowOffsetX                   |  Integer  | *set the X offset value of the shadow*                       |
| ShadowOffsetY                   |  Integer  | *set the Y offset value of the shadow*                       |
| ShadowSize                      |  Integer  | *set the width of the shadow of the control*                 |
| :star:Value               |  Boolean  | *set/return a boolean value when* OptionBehavior *is enabled* |
| :star:Version             |  String   | *return the version number of the control*                   |



## Events

:star:: additions versus original LabelPlus

| **Events**                                                   | **Description**                                              |
| ------------------------------------------------------------ | ------------------------------------------------------------ |
| Click()                                                      | *raised  by clicking on the control*                         |
| DblClick()                                                   | *raised  by double clicking on the control*                  |
| :star:Change()                                               | *raised  when captions are modified*                         |
| :star:ChangeValue(Value As Boolean)                          | *raised  when* OptionBehavior *is enabled and return* Value *property value. (True or False)* |
| :star:CrossClick()                                           | *raised  when* CrossVisible *is enabled and cross is clicked* |
| MouseDown(Button As Integer, Shift As Integer, X As Single, Y  As Single) | *raised  when a mouse button is held down over the control, and return mouse button  pressed, special key pressed and coordinates X and Y* |
| MouseMove(Button As Integer, Shift As Integer, X As Single, Y  As Single) | *raised  when a mouse button is held down while moving it over the control, and return  mouse button pressed, special key pressed and coordinates X and Y* |
| MouseUp(Button As Integer, Shift As Integer, X As Single, Y As  Single) | *raised  when the mouse button pressed over the control is released, and return mouse  button pressed, special key pressed and coordinates X and Y* |
| MouseEnter()                                                 | *raised  when the mouse pointer enters the control area*     |
| MouseLeave()                                                 | *raised  when the mouse pointer left the control area*       |
| MouseOver()                                                  | *rises  when the mouse pointer hovers over the control area* |
| PrePaint(hdc As Long, X As Long, Y As  Long)                 | *this event is fired before the paint of the control on the form/container* |
| PostPaint(ByVal hdc As Long)                                 | *this event is fired after the paint of the control on the form/container, especially useful for painting aditional elements on the control with the* Drawtext, DrawLine *and* Polygon *functions* |
| PictureDownloadProgress(BytesMax As  Long, BytesLeidos As Long) | *this event is raised while image loading (loaded with functions* PictureFromStream *or* PictureFromStream*) is performed in the control* |
| PictureDownloadComplete()                                    | *raised when the image loading is completed*                 |
| PictureDownloadError()                                       | *raised if the image loading get an error*                   |

## Functions

:star:: additions versus original LabelPlus

| Functions                                                    | Description                                                  |
| ------------------------------------------------------------ | ------------------------------------------------------------ |
| **DrawLine(**hdc, X, Y1, X2, Y2, [oColor], [Opacity], [PenWidth]**)** | *allows to draw a line over the control*                     |
| **DrawText(**hdc, Text, X, Y, Width, Height, Font, ForeColor, [ColorOpacity], [HorizontalAlign], [VerticalAlign], [WordWrap]**)** | *allows to draw text over the control*                       |
| :star:IsMouseInExtender                                      | *this function returns TRUE or FALSE whether the mouse pointer is in the control or not. use this instead of* MouseEnter, MouseOver *or* MouseLeave *when you can't or don't want to depend on those events* |
| :star:**LoadImagefromPath(**path_to_file**)**                | *Load an image on the control from path string*              |
| PictureDelete                                                | *delete the image from the control*                          |
| **PictureFromStream(**array_of_bytes**)**                    | *load an image on the conrol from an array of bytes (internally the control uses this function to load an image from the properties page)* |
| **PictureFromURL(**sUrl, [UseCache], [DrawProgress]**)**     | *Load an image on the control from url*                      |
| **Polygon(**hDc, PenWidth, Color, Opacity, ParamArray vPoints()**)** | allows you to draw polygons on the control                   |
| Refresh                                                      | *Refresh the drawing of the control*                         |



## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See deployment for notes on how to deploy the project on a live system.

### Prerequisites



GDI Plus, use from Win7 ahead...

### Installing

Copy Files to your Project folder, include/reference to this and set this usercontrol to Private

```
axLabelPlus.ctl    <UserControl>
axLabelPlus.ctx    <resources of Usercontrol>
axLabelPlus.pag    <Property Page of Usercontrol>
axLabelPlus.pgx    <resources of Property Page>
```

Or Compile the usercontrol to OCX (ActiveX), and reference to your VB6 Project

```
AXLPCTRL.OCX
```

...

## Built With

- *Clasic and Beloved* **Visual Basic 6 - ServicePack 6**  (Visual Basic *is trademark of* ![micrologo](https://user-images.githubusercontent.com/61160830/119084996-d1da1a80-b9d0-11eb-9132-f0f3e062f9d0.jpg))


## Versioning

We use [SemVer](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/your/project/tags).

## Authors

- **Leandro Ascierto** - *Initial work* - [Leandro Ascierto VB6 Latin Blog & Forums](http://leandroascierto.com/blog/)
- **AxioUK** - David Rojas *Editor & Forum User* - [Leandro Ascierto VB6 Latin Blog & Forums](http://leandroascierto.com/blog/)

## License

This project is free to use, modify and sharing... only mention the authors in the credits :smile:

