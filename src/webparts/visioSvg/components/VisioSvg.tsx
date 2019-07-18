import * as React from 'react';
import styles from './VisioSvg.module.scss';
import { IVisioSvgProps } from './IVisioSvgProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import * as Snap from 'snapsvg-cjs';
import { autobind } from 'office-ui-fabric-react';
import * as svgPanZoom from 'svg-pan-zoom';


export interface IVisioSvgState {
  title: string;
  comment: string;
  output: any;
}
export default class VisioSvg extends React.Component<IVisioSvgProps, IVisioSvgState> {
  constructor(props: IVisioSvgProps) {
    super(props);
    console.log("showVisioSVG");
    //  this.callback = Guid.create();
    this.state = {
      title: "",
      comment: "",
      output: "-"
    };
  }
  private svgContainer: any;
  // private svgObject: any;
  private svgPaper: any;
  private rootPaper: any;
  private groupcontext: string = "v:groupContext";
  private selectionRect: any;
  private hoveredShape: any;
  private panZoom: any;



  public render(): React.ReactElement<IVisioSvgProps> {
    return (
      <div>
        <table>
          <tbody>
            <tr>
              <td><div id="title">{this.state.title}</div>
                <div id="comment">{this.state.comment}</div>
                <div id="output">
                  {this.state.output}
                </div>
                <div className={styles.visioSvg}>
                  <div id='some-empty-div-on-the-page'></div>
                </div>
              </td>
              <td>
              </td></tr>
          </tbody>
        </table>
      </div>);
  }
  @autobind
  public componentDidMount() {
    //   if (this.props.context && this.props.filter) {
    //     SetContext(this.props.context, this.props.filter);
    // var selectionRect, hoveredShape;

    // Initialize Snap.svg and tell it to use the SVG data in the <object> tag
    // var rootPaper = Snap("#svg-object");

    // Instead of using <object> tags, you can also load SVG files using Snap.svg itself. Here is an example:

    this.svgContainer = Snap("#some-empty-div-on-the-page");
    let me: VisioSvg = this;

    Snap.load("https://semtalk.sharepoint.com/teams/tender/SVG/tender.svg", (data: any) => {
      //    Snap.load("./VisioFlowchart.svg", (data: any) => {
      this.svgContainer.append(data);
      this.svgPaper = Snap(this.svgContainer.node.firstElementChild);
      this.rootPaper = this.svgPaper.paper;
      this.groupcontext = "v:groupcontext";
      this.selectionRect = this.rootPaper.rect(0, 0, 0, 0);



      console.log(this.panZoom);
      console.log(this.rootPaper);
      this.setState({
        title: this.rootPaper.select("title").node.textContent,
        // comment: this.rootPaper.select("desc").node.textContent
      });
      this.rootPaper.selectAll("g").forEach((elm: any, _i) => {
        // Browsers will pick up the <title> tag and display as a tooltip. This can interfere with your own application, so we can remove it this way.
        // But the title tells us what kind of Visio shape it is, so we keep a copy of it in an atribute instead.
        if (elm.select("title") && elm.select("title").node) {
          var type = elm.select("title").node.textContent;
          elm.attr("shapeType", type);
          elm.select("title").remove();
        }
      });
      // this.rootPaper.selectAll("g[id^=shape]").forEach((elm, _i) => {
      this.rootPaper.selectAll("g").forEach((elm, _i) => {
        // Click event
        elm.click((evt) => {

          // Call a helper function (see further down) to get the Visio shape clicked
          var shape = me.parseShapeElement(elm);
          if (!shape) return;

          // console.log("SVG Click");

          // Clear the previous selection box (if any)
          // if (me.selectionRect) {
          //   me.selectionRect.remove();
          // }

          // Draw a box around the selected shape
          // Unfortunately this is not as precise as one is used to when working with CSS, so the box might be slightly off
          // One reasone for this is that getBBox() does not account for strokeWidth, and I have yet to find a way to read this (especially in more complex shape-groups).
          //  me.selectionRect = this.rootPaper.rect(shape.x, shape.y, shape.width, shape.height);

          // The new selection rect will not be visible unless we set some SVG style attributes.
          // These are similar to HTML/CSS attributes, sometimes even idential.
          // Note that we are setting more than one attribute in a single call here.
          // "fill" should be evident. "stroke" refers to the border around an SVG element.
          me.selectionRect.attr({
            x: shape.x,
            y: shape.y,
            width: shape.width,
            height: shape.height,
            fill: "none",
            stroke: "red",
            strokeWidth: 1
          });

          // Setting "pointerEvents" to none means that the mouse will never be able to click/hover this element.
          me.selectionRect.attr({
            pointerEvents: "none"
          });

          // Setting the "shapeRendering" attribute allows to hint to the browser how shapes should rendered.
          // In this example, "geometricPrecision" appears to give best result. Also try "optimizeSpeed" and "crispEdges"
          me.selectionRect.attr({
            shapeRendering: "geometricPrecision"
          });

          // Finally stop the click event from bubbling up to the Visio "page"
          evt.preventDefault();
          evt.stopPropagation();
        });
        elm.mouseover((evt) => {

          var shape = me.parseShapeElement(elm);
          if (!shape) return;

          // Set cursor to pointer to indicate that we can click on shapes
          // (We could have set this elsewhere; this was just a convenient place)
          elm.attr({
            cursor: "pointer"
          });

          // When hovering a shape we want some kind of indication.
          // Setting the shape opacity to 50% works fine in our example. Depending on the background underneath the shape this may not work though.
          // Alternatively we could have used a something similar to when selecting shapes, drawing a rectangle around it.
          // Ideally we would want to clone the shape and display on top. Unfortunately I have not gottent this to work properly. The problem likely stems from Visio shapes being complex elements put together of multiple parts.
          if (me.hoveredShape) {
            // First reset any previous shape that we hovered
            me.hoveredShape.attr({
              fillOpacity: "1",
              strokeOpacity: "1"
            });
          }
          // Set the opacity attributes to 50% on the current hovered shape
          me.hoveredShape = elm;
          elm["attr"]({
            fillOpacity: "0.5",
            strokeOpacity: "0.5"
          });

          // Print data found in the Visio shape to the output panel
          // To see how these were obtained, see parseShapeElement() below
          if (shape) {
            me.setState({
              output: shape.text + "(" + shape.id + ")",
            });
          } else {
            me.setState({
              output: "-",
            });
          }
        });
        elm.mouseout((evt) => {
          // Reset a few things when the mouse is no longer hovering the shape
          //  $("#output").empty();
          me.setState({
            output: "-",
          });
          if (me.hoveredShape) {
            me.hoveredShape.attr({
              fillOpacity: "1",
              strokeOpacity: "1"
            });
            me.hoveredShape = null;
          }
        });

      });
      // Also add a click event handler to the background
      this.rootPaper.click((evt) => {
        console.log("SVG Click on Paper");
        if (me.selectionRect) {
          me.selectionRect.attr({
            x: 0,
            y: 0,
            width: 0,
            height: 0,
            fill: "none",
            stroke: "red",
            strokeWidth: 0
          });
          // this.selectionRect.remove();
        }
      });
      this.resizeSvgAndContainer(this.svgContainer, this.rootPaper);

      this.panZoom = svgPanZoom(this.svgPaper.node, {
        zoomEnabled: true,
        controlIconsEnabled: true
      });
    });

  }

  // This helper function will take an element and try to parse it as a Visio shape
  // It will return a new object with the properties we are interested in
  private parseShapeElement(elm: any): any {

    // Figuring out where things are located in Visio SVG:s and what they are named, such as "v:groupContext" was done by inspecting examples in and ordinary text editor
    var elementType = elm.node.attributes[this.groupcontext].value;
    if ((elementType === "shape" || elementType === "group") && elm.node != null) {
      //  if (elm.node != null) {

      // Create the object to hold all data we collect
      var shape = {};

      // The shape type tells us what kind of Visio shape we are dealing with
      shape["type"] = elm.node.attributes["shapeType"].value;

      let binteresting: boolean=false;

      // Make sure this Visio shape is of interest to us.
      // "dynamic connector" corresponds to arrows
      // "sheet" can be the background or container objects
      /*       var doNotProcess = ["sheet", "dynamic connector"];
            var type = shape["type"].toLowerCase();
            for (let i = 0; i < doNotProcess.length; i++) {
              if (type.indexOf(doNotProcess[i]) !== -1) {
                return null;
              }
            } */

      // Let begin collecting data!
      shape["paper"] = elm.paper;

      // Each shape has a unique id
      shape["id"] = elm.node.attributes["id"].value;

      if (elm.node.attributes["v:mid"]) {
        shape["shapeid"] = "Sheet." + elm.node.attributes["v:mid"].value;
      }

      // Shape position is relative to the SVG coordinates, not the screen/browser
      shape["x"] = elm.getBBox().x;
      shape["y"] = elm.getBBox().y;
      shape["width"] = elm.getBBox().width;
      shape["height"] = elm.getBBox().height;

      // Get the text inside the shape
      shape["text"] = "";
      if (elm.select("desc") && elm.select("desc").node) {
        shape["text"] = elm.select("desc").node.textContent;
      }
      shape["props"] = {};
      shape["defs"] = {};

      let pn = elm.node.parentNode;
      if (pn.attributes["xlink:href"]) {
        shape["xlink_href"] = pn.attributes["xlink:href"].value;
      }
      if (pn.attributes["xlink:show"]) {
        shape["xlink_show"] = pn.attributes["xlink:show"].value;
      }
      if (pn.attributes["xlink:title"]) {
        shape["xlink_title"] = pn.attributes["xlink:title"].value;
      }
      if (pn.attributes["target"]) {
        shape["xlink_target"] = pn.attributes["target"].value;
      }

      let cl = elm.node.childNodes;
      for (let j = 0; j < cl.length; j++) {
        let cn = cl[j];
        if (cn.nodeName == "v:custprops") {
          let cpl = cn.childNodes;
          for (let k = 0; k < cpl.length; k++) {
            let un = cpl[k];
            if (un.tagName == "v:cp") {
              if (un.attributes["v:nameu"] && un.attributes["v:val"]) {
                shape["defs"][un.attributes["v:nameu"].value] = un.attributes["v:val"].value;
              }
            }
          }
        }
        if (cn.nodeName == "v:userdefs") {
          let ul = cn.childNodes;
          for (let k = 0; k < ul.length; k++) {
            let un = ul[k];
            if (un.nodeName == "v:ud") {
              if (un.attributes["v:nameu"] && un.attributes["v:val"]) {
                let nameu: string= un.attributes["v:nameu"].value;
                if (nameu.toLowerCase()=="SemTalkInstID".toLowerCase()) {
                  binteresting=true;
                }
                shape["defs"][nameu] = un.attributes["v:val"].value;
              }
            }
          }
        }
      }
      if (!binteresting) {
        return null;
      }
      // console.log(shape);
      return shape;
    } else {
      // Not a Visio shape
      return null;
    }
  }


  private resizeSvgAndContainer(_objectElement: any, rp: any) {
    return;
    /*     // Get bounding box of the (Visio) page
        this.rootPaper.attr({ width: "400px", height: "800px" });
        rp.selectAll("g").forEach((elm: any) => {
          if (elm.node.attributes[this.groupcontext] && elm.node.attributes[this.groupcontext].value === "foregroundPage") {

            var visioPage = elm.node;
            // The "Bounding Box" contains information about an SVG element's position
            var bbox = visioPage.getBBox();
            var x = bbox.x;
            var y = bbox.y;
            var w = bbox.width;
            var h = bbox.height;

            // Figure out a new viewBox that shows as much as possible of the drawing
            // The viewbox is a property of SVG that specifies what part of the drawing to display. The actual drawing can extend beyong this viewbox so you would have to pan the drawing to view it all.
            // It is not perfect. This is probably because getBBox does not include account for strokeWidth, and I have yet to find out a way to figure this out.
            // This can cause shapes to be clipped. I am adding marging to the new viewbox to try and fix this. Using a relative marginY appears to give decent result for most of my needs. You may need to try something different yourself.
            var marginX = 1;
            var marginY = (w / h);
            var newViewBox = (x - marginX) + " " + (y - marginY) + " " + (w + marginX * 2) + " " + (h + marginY * 2);

            // The width of the container is fixed. But we can alter the height to show as much as possible of the drawing at its specified aspect ratio.
            //  if (host != undefined) {
            // let hw: number = host.width() as number;
            // host.height(hw / (w / h));
            //  host.height(300);
            //  host.width(300);
            // }
            // Set SVG to make Visio page fill entire object canvas
            // Here I am also using Snap's animation feature for a nice effect when loading the page
            // 300 is the animation duration, "mina.easeinout" is the animation easing. Other easings are: easeinout, linear, easein, easeout, backin, backout, elastic & bounce.
            // rp.animate({ viewBox: newViewBox }, 300, mina.easeinout);
            this.rootPaper.animate({ viewBox: newViewBox }, 300);
          }
        }); */
  }

}
