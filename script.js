$.fn.extend({
  printMatrix: function (config) {
    let matrix = config.matrix;
    let hasAxis =
      config.hasAxis === undefined || typeof config.hasAxis !== "boolean"
        ? true
        : config.hasAxis;
    let title = config.title === undefined ? "" : config.title;
    if (matrix == undefined) return;
    let height = config.height === undefined ? "180px" : config.height;
    let html =
      '<table class="ui definition unstackable celled table center aligned matrix selectable">';
    if (hasAxis) {
      html += "<thead><tr><th>" + title + "</th>";
      numberCol = 0;
      matrix.forEach((element) => {
        numberCol = element.length > numberCol ? element.length : numberCol;
      });

      for (let i = 1; i <= numberCol; i++) {
        html += "<th>" + i + "</th>";
      }
      html += "</tr></thead>";
    }
    html += "<tbody>";
    matrix.forEach((element, index) => {
      html += "<tr>";
      if (hasAxis) {
        html += "<td>" + (index + 1) + "</td>";
      }
      element.forEach((subElement) => {
        html += "<td>" + subElement + "</td>";
      });
      html += "</tr>";
    });
    html += "</tbody></table>";
    this.html(html);
    let table = this;
    var max = 0;
    $("table td", this).each(function () {
      max = Math.max($(this).outerWidth(true), max);
    });
    $("table th", this).each(function () {
      max = Math.max($(this).outerWidth(true), max);
    });
    $("tbody", this).css("height", height);
    $("table td", this).css("min-width", max);
    $("table th", this).css("min-width", max);
    $("thead th:nth-child(1)", this).css(
      "min-height",
      $("thead th:nth-child(2)", this).outerHeight(true)
    );
    $("tbody", this).scroll(function (e) {
      let left = $(this).scrollLeft();
      $("thead", table).css("left", -left);
      $("thead th:nth-child(1)", table).css("left", left);
      $("tbody td:nth-child(1)", table).css("left", left);
    });
  },
});

window.onload = function () {
  $(".ui.sticky").sticky({
    offset: 50,
  });

  let firstTime = true;
  $("#r-slider").slider({
    min: 0,
    max: 1,
    start: 0.3,
    step: 0.01,
    onMove: function (val) {
      $("#r-value").html(val);
      R = val;
      if (!firstTime) setup();
    },
  });
  $("#alpha-slider").slider({
    min: 0,
    max: 1,
    start: 1,
    step: 0.01,
    onMove: function (val) {
      $("#alpha-value").html(val);
      alpha = val;
      if (!firstTime) setup();
    },
  });

  $("#umin-slider").slider({
    min: 0,
    max: 1,
    start: 1,
    step: 0.01,
    onMove: function (val) {
      $("#umin-value").html(val);
      u_min = val;
      if (!firstTime) setup();
    },
  });

  firstTime = false;

  $(".team").click(function () {
    $("#team-info-modal").modal("show");
  });
  $("#close-info-modal").click(function () {
    $("#team-info-modal").modal("hide");
  });
  $("#problem").click(function () {
    $("#problem-modal").modal("show");
  });
  $("#close-problem-modal").click(function () {
    $("#problem-modal").modal("hide");
  });
  $("#reload").click(async function () {
    $("#loading").show();
    $("#topology").empty();
    $("#reload").addClass("loading");
    await reload(false);
    $("#reload").removeClass("loading");
  });
  $(":checkbox").change(function () {
    DRAW_CONVEX_HULL = $("#ck1").prop("checked");
    DRAW_MAX_DISTANCE = $("#ck2").prop("checked");
    DRAW_BACKBONE_CIRCE = $("#ck3").prop("checked");
    DRAW_LINE_FROM_BACKBONE_TO_ACCESS = $("#ck4").prop("checked");
    MENTOR_1 = $("#ck5").prop("checked");
    MENTOR_2 = $("#ck6").prop("checked");
    DIJSKTRA_TRAFFIC_ONLY = $("#dijsktra-traffic-only").prop("checked");
    HIDE_NO_TRAFFIC = $("#hide-no-traffic").prop("checked");
    setup();
  });
};

function reload(init) {
 
  if (init) {
    let w,T;
    var mathDiv = document.getElementById("math");
    var displayDiv = document.getElementById("display");
  }

  w = new Array(nodeNumber);
  for (let i = 0; i < nodeNumber; i++) {
    let result =0;
    if(i < 3 || i >= 87){
      result += 1;
    }
    else{
      result +=2;
    }
    if(i < 23 || i>=67){
      result += 2;
    }
    if(i < 5 || i >= 85){
      result += 5;
    }
    if(i == 6 || i == 27){
      result += 5;
    }
    if(i == 26 || i ==46){
      result += 10;
    }
    if(i == 11 || i == 45){
      result += 26;
    }
    if(i == 5 || i == 87){
      result += 24;
    }
    w[i] = result;
  }
  

  

  // Gán lưu lượng giữa các nút
  T = new Array(nodeNumber);
  for (let i = 0; i < T.length; i++) {
    T[i] = new Array(nodeNumber).fill(0);
  }
  T[7][28] = 5;
  T[28][7] = 5;
  T[26][47] = 10;
  T[47][26] = 10;
  T[12][46] = 26;
  T[46][12] = 26
  T[6][88] = 24;
  T[88][6] = 24;
  for (let i = 0; i < nodeNumber; i++) {
    for (let j = 0; j < nodeNumber; j++) {
      if (T[i][j] == 0) {
        if (j === i + 3) {
          T[i][j] += 1;
          T[j][i] += 1;
        }
        if (j === i + 67) {
          T[i][j] += 2;
          T[j][i] += 2;
        }
        if( j == i + 85){
          T[i][j] += 5;
          T[j][i] += 5;
        }
      }
    }
  }

  // $("#flow").printMatrix({
  //   title: "T",
  //   matrix: T,
  //   height: "300px",
  // });

  setNodes();

  return new Promise((resolve) =>
    setTimeout(() => {
      if(!init) setup();
      resolve();
    }, 1000)
  );
}

function getRandomInt(min, max) {
  min = Math.ceil(min);
  max = Math.floor(max);
  return Math.floor(Math.random() * (max - min) + min); //The maximum is exclusive and the minimum is inclusive
}

var DRAW_CONVEX_HULL = $("#ck1").prop("checked");
let DRAW_MAX_DISTANCE = $("#ck2").prop("checked");
let DRAW_BACKBONE_CIRCE = $("#ck3").prop("checked");
let DRAW_LINE_FROM_BACKBONE_TO_ACCESS = $("#ck4").prop("checked");
let MENTOR_1 = $("#ck5").prop("checked");
let MENTOR_2 = $("#ck6").prop("checked");
let DIJSKTRA_TRAFFIC_ONLY = $("#dijsktra-traffic-only").prop("checked");
let HIDE_NO_TRAFFIC = $("#hide-no-traffic").prop("checked");
let x_matrix = 1200;
let y_matrix = 1200;
let nodeNumber = 90;

let nodes = [];
let W = 2;
let R = 0.3;
let C = 15;
let alpha = 1;
let u_min = 0.8;
const wNumberElement = document.getElementById("number-w");
wNumberElement.addEventListener("change", (e) => {
  W = parseInt(wNumberElement.value);
  setup();
});

const nodeNumberElement = document.getElementById("number-nodes");
nodeNumberElement.addEventListener("change", (e) => {
  nodeNumber = parseInt(nodeNumberElement.value);
});

nodeNumberElement.addEventListener("keyup", function (event) {
  if (event.key === "Enter") {
    document.getElementById("reload").click();
  }
});

reload(true);

function setNodes() {
  nodes = [];
  for (let i = 0; i < nodeNumber; i++) {
    let point = {
      x: Math.floor(Math.random() * x_matrix),
      y: Math.floor(Math.random() * y_matrix),
      index: i,
      classified: false,
    };
    nodes[i] = point;
  }
}

function lineNode(a, b, color = 255) {
  stroke(color);
  strokeWeight(1.5);
  line(
    Math.round(a.x / 2),
    Math.round(a.y / 2),
    Math.round(b.x / 2),
    Math.round(b.y / 2)
  );
}

function lineBlackboneNode(a, b, color = 255) {
  stroke(color);
  strokeWeight(2);
  line(
    Math.round(a.x / 2),
    Math.round(a.y / 2),
    Math.round(b.x / 2),
    Math.round(b.y / 2)
  );
}

function drawBackboneCircle(e, R_backbone, color = 255) {
  noFill();
  stroke(color);
  circle(Math.round(e.x / 2), Math.round(e.y / 2), R_backbone);
}

function drawBackboneNode(e, color = 255) {
  fill(color);
  stroke(color);
  rect(Math.round(e.x / 2) - 3.5, Math.round(e.y / 2) - 3.5, 8, 8);
}

function drawCenterNode(e, color = 255) {
  fill(color);
  stroke(color);
  rect(Math.round(e.x / 2) - 3.5, Math.round(e.y / 2) - 3.5, 11, 11);
}

// Kiểm tra a->b->c có phải theo chiều kim đồng hồ không bằng cách kiểm tra xem
// b có nằm trên nửa mặt phẳng dương có bờ là đường thẳng ac không (do a, b, c
// đã sắp xếp theo chiều tăng trục hoành)
function isClockwise(a, b, c) {
  return a.x * (c.y - b.y) + b.x * (a.y - c.y) + c.x * (b.y - a.y) > 0;
}
// Kiểm tra a->b->c có phải ngược chiều kim đồng hồ không
function isAntiClockwise(a, b, c) {
  return a.x * (c.y - b.y) + b.x * (a.y - c.y) + c.x * (b.y - a.y) < 0;
}

function isNodeInBackboneCircle(checking_node, backbone_node, R_backbone) {
  return  calcDistance(checking_node, backbone_node) <= R_backbone;
}

function removeNodeFromBackbone(node, backbone_node) {
  for (let i = 0; i < backbone_node.access.length; i++) {
    if (backbone_node.access[i] === node) {
      backbone_node.access.splice(i, 1);
    }
  }
}

function calcDistance(node1, node2) {
  return Math.sqrt(
    Math.pow(node1.x - node2.x, 2) + Math.pow(node1.y - node2.y, 2)
  );
}

function setup() {
  $("#loading").hide();
  let width = x_matrix + 5;
  let height = y_matrix + 5;

  let myCanvas = createCanvas(Math.round(width / 2), Math.round(height / 2));
  myCanvas.parent("topology");
  background(0);

  // Vẽ nodes
  let points = JSON.parse(JSON.stringify(nodes));

  points.forEach((point) => {
    //point(Math.round(point.x/2), Math.round(point.y/2));
    circle(Math.round(point.x / 2), Math.round(point.y / 2), 4);
  });

  // Thêm trọng số vào mỗi nút
  points.forEach((element, index) => {
    element.w = w[index];
  });

  // Tìm đường bao (phục vụ cho việc tìm khoảng cách xa nhất)
  let sorted = [...points].sort(function (a, b) {
    return a.x - b.x || (a.x == b.x && a.y - b.y);
  });
  let p1 = sorted[0],
    p2 = sorted[nodeNumber - 1];
  let up = [],
    down = [],
    convex_hull = [];
  up.push(p1);
  down.push(p1);
  for (let i = 0; i < nodeNumber; i++) {
    if (i == nodeNumber - 1 || isClockwise(p1, sorted[i], p2)) {
      while (
        up.length >= 2 &&
        !isClockwise(up[up.length - 2], up[up.length - 1], sorted[i])
      )
        up.pop();
      up.push(sorted[i]);
    }
    if (i == nodeNumber - 1 || isAntiClockwise(p1, sorted[i], p2)) {
      while (down.length >= 2 && !isAntiClockwise(down[1], down[0], sorted[i]))
        down.shift();
      down.unshift(sorted[i]);
    }
  }
  up.pop();
  down.pop();
  convex_hull = convex_hull.concat(up);
  convex_hull = convex_hull.concat(down);
  if (DRAW_CONVEX_HULL) {
    for (let i = 0; i < convex_hull.length; i++) {
      lineNode(convex_hull[i], convex_hull[(i + 1) % convex_hull.length]);
    }
  }

  // Tìm khoảng cách xa nhất
  let max_distance = {
    node1: 0,
    node2: 0,
    distance: 0,
  };

  for (let i = 0; i < convex_hull.length - 1; i++) {
    for (let j = i + 1; j < convex_hull.length; j++) {
      let distance = calcDistance(convex_hull[i], convex_hull[j]);
      if (distance > max_distance.distance) {
        max_distance.node1 = convex_hull[i];
        max_distance.node2 = convex_hull[j];
        max_distance.distance = distance;
      }
    }
  }
  let R_backbone = R * max_distance.distance;

  if (DRAW_MAX_DISTANCE) {
    lineNode(max_distance.node1, max_distance.node2, color(255, 0, 0));
  }

  // Phân loại nút
  let backbones = [];
  points.forEach(function (e, i) {
    if (w[i] / C > W) {
     
      backbones.push(e);
      e.is_backbone = true;
      e.is_access = false;
      e.classified = true;
      // drawBackboneNode(e);
      if (DRAW_BACKBONE_CIRCE) {
        drawBackboneCircle(e, R_backbone, 100);
      }
    }
  });

  while (true) {
    
    let x_numerator = 0,
      y_numerator = 0,
      denominator = 0;
    points.forEach((e) => {
      if (e.is_backbone != true) {
        // Nút chưa được phân loại và nút truy nhập
        backbones.forEach(function (b) {
          
          if (isNodeInBackboneCircle(e, b, R_backbone)) {
            // Nếu nút cần kiểm tra nằm trong nút backbone b
            e.is_access = true; // Chuyển thành nút truy nhập
            e.is_backbone = false;
            e.classified = true;
            let localDistance = calcDistance(e, b); // Tính khoảng cách giữa nút này và nút backbone
            if (
              e.backbone_distance === undefined ||
              e.backbone_distance > localDistance
            ) {
              if (e.backbone !== undefined)
                removeNodeFromBackbone(e, e.backbone);
              if (b.access === undefined) b.access = [];
              b.access.push(e);
              e.backbone = b; // Nếu khoảng cách mới nhỏ hơn khoảng cách cũ thì nút truy nhập sẽ thuộc cây backbone mới
              e.backbone_distance = calcDistance(e, b);
            }
          }
        });
      } else {
        if (e.access === undefined) e.access = [];
      }
    });
    
   
    let dc = [];
    let dcmax = 0,
      w_at_dcmax = 0,
      untype = 0;
    points.forEach((e) => {
      if (!e.classified) {
        console.log(e);
        untype++;
        x_numerator += e.x * e.w;
        y_numerator += e.y * e.w;
        denominator += e.w;
      }
    });
    let centerNode = {
      x: x_numerator / denominator,
      y: y_numerator / denominator,
    };
    if (untype == 0) {
      break;
    }
    points.forEach((e, i) => {
      if (!e.classified) {
        console.log(centerNode);
        dc[i] = calcDistance(e, centerNode);
        if (dc[i] >= dcmax) {
          dcmax = calcDistance(e, centerNode);
          w_at_dcmax = w[i];
        }
      }
    });
  
    let merit_max = 0;
    index_at_merit_max = 0;
    points.forEach(function (e, i) {
      if (!e.classified) {
        let merit;
        if (dcmax != 0) {
          merit = (0.5 * (dcmax - dc[i])) / dcmax + (0.5 * e.w) / w_at_dcmax;
        } else {
          merit = (0.5 * e.w) / w_at_dcmax;
        }
        if (merit >= merit_max) {
          merit_max = merit;
          index_at_merit_max = i;
        }
      }
    });
    backbones.push(points[index_at_merit_max]);
    points[index_at_merit_max].is_access = false;
    points[index_at_merit_max].is_backbone = true;
    points[index_at_merit_max].classified = true;
    // drawBackboneNode(points[index_at_merit_max]);
    if (DRAW_BACKBONE_CIRCE)
      drawBackboneCircle(points[index_at_merit_max], R_backbone, 100);
  }
  if (DRAW_LINE_FROM_BACKBONE_TO_ACCESS) {
    points.forEach((e) => {
      if (e.is_access) {
        lineNode(e, e.backbone);
      }
    });
  }

  // In ra bảng nodes
  let htmlNodesTable =
    '<table class="ui celled table center aligned unstackable selectable"><thead><tr><th>Nút</th><th>Nút xương sống</th><th>Nút truy nhập</th>';
  htmlNodesTable += "</tr></thead><tbody>";
  points.forEach((element, index) => {
    htmlNodesTable += "<tr><td>" + (index + 1) + "</td>";
    htmlNodesTable +=
      "<td>" +
      (element.is_backbone ? '<i class="check icon"></i>' : "") +
      "</td>";
    htmlNodesTable +=
      "<td>" +
      (element.is_access ? '<i class="check icon"></i>' : "") +
      "</td></tr>";
  });
  htmlNodesTable += "</tbody></table>";
  $("#nodes").html(htmlNodesTable);
  $("#nodes").scroll(function () {
    if ($(this).scrollTop() == 0) {
      $(this).css("border-top", "0");
    } else {
      $(this).css("border-top", "1px solid rgba(34,36,38,.1)");
    }
  });

  // Tính lưu lượng giữa các nút backbone
  let numberBackbones = backbones.length;

  backbones.sort(function (a, b) {
    return a.index - b.index;
  });

 

  let T_b = [];
  backbones.forEach((b1, i1) => {
    T_b[i1] = [];
    backbones.forEach((b2, i2) => {
      let trafic = 0;
      if (i1 !== i2) {
        b1.access.forEach((a1) => {
          trafic += T[a1.index][b2.index];
          b2.access.forEach((a2) => {
           
            trafic += T[a1.index][a2.index];
          });
        });
        b2.access.forEach((a) => {
          trafic += T[a.index][b1.index];
        });
        trafic += T[b1.index][b2.index];
      }
      T_b[i1][i2] = trafic;
    });
  });

  for (let i = 0; i < numberBackbones - 1; i++) {
    for (let j = i + 1; j < numberBackbones; j++) {
      if (T_b[i][j] !== 0) {
       
        
      }
    }
  }

  $("#backbone-flow").printMatrix({
    title: "T<sub>b</sub>",
    matrix: T_b,
    height: "300px",
  });

  // 3. MENTOR
  // 3.1. Tìm nút backbone trung tâm
  // Tính moment của các nút backbone, tìm nút backbone trung tâm (có moment nhỏ nhất)
  let min_moment_index = backbones[0].index;
  let sourceBackboneIndex;
  backbones.forEach((e1, i1) => {
    e1.moment = 0;
    backbones.forEach((e2) => {
      if (e1.index != e2.index) {
        e1.moment += calcDistance(e1, e2) * e2.w;
      }
    });
    if (e1.moment <= points[min_moment_index].moment) {
      min_moment_index = e1.index;
      sourceBackboneIndex = i1;
    }
  });
  // drawCenterNode(points[min_moment_index], color(0,100,0));

  // In danh sách nút backbones
  let htmlBackbonesTable =
    '<table class="ui celled table center aligned unstackable selectable"><thead><tr><th>Nút backbones</th><th>Nút ban đầu</th>';
  htmlBackbonesTable += "</tr></thead><tbody>";
  backbones.forEach((element, index) => {
    htmlBackbonesTable +=
      "<tr" +
      (index == sourceBackboneIndex ? ' class="warning"' : "") +
      "><td>" +
      (index + 1) +
      (index == sourceBackboneIndex ? " (nút trung tâm)" : "") +
      "</td>";
    htmlBackbonesTable += "<td>" + (element.index + 1) + "</td></tr>";
  });
  htmlBackbonesTable += "</tbody></table>";
  $("#backbones").html(htmlBackbonesTable);
  $("#backbones").scroll(function () {
    if ($(this).scrollTop() == 0) {
      $(this).css("border-top", "0");
    } else {
      $(this).css("border-top", "1px solid rgba(34,36,38,.1)");
    }
  });

  // 3.2. Xây dựng cây Prim - Dijsktra cho các nút backbone
 
  let label = new Array(numberBackbones).fill(Infinity); // lable[i] là nhãn của nút i
  let S = new Array(numberBackbones).fill(false); // S == true nghĩa là điểm đã được xét
  let P = new Array(numberBackbones).fill(sourceBackboneIndex); // Hàm tiền bối
  label[sourceBackboneIndex] = 0;
  let labelMinIndex = sourceBackboneIndex;
  for (let i = 0; i < numberBackbones; i++) {
    let newLabelMinIndex,
      newLabelMinValue = Infinity;
    S[labelMinIndex] = true;
    for (let j = 0; j < numberBackbones; j++) {
      if (S[j] == false) {
        if (
          (!DIJSKTRA_TRAFFIC_ONLY || T_b[labelMinIndex][j] > 0) &&
          alpha * label[labelMinIndex] +
            calcDistance(backbones[labelMinIndex], backbones[j]) <
            label[j]
        ) {
          label[j] =
            alpha * label[labelMinIndex] +
            calcDistance(backbones[labelMinIndex], backbones[j]);
          P[j] = labelMinIndex;
        }
        if (label[j] < newLabelMinValue) {
          newLabelMinValue = label[j];
          newLabelMinIndex = j;
        }
      }
    }
    labelMinIndex = newLabelMinIndex;
  }

  let links = new Array(numberBackbones);
  for (let i = 0; i < numberBackbones; i++) {
    links[i] = new Array(numberBackbones).fill({
      n: 0,
      u: 0,
      d: 0,
      trafic: 0,
    });
  }

  let linksTraffic = new Array(numberBackbones);
  for (let i = 0; i < numberBackbones; i++) {
    linksTraffic[i] = new Array(numberBackbones).fill(0);
  }

  // Tìm ma trận hops và ma trận đường đi path
  let hop = [];
  let path = [];
  for (let i = 0; i < numberBackbones; i++) {
    hop[i] = [];
    path[i] = [];
    for (let j = 0; j < numberBackbones; j++) {
      let ci = 0,
        cj = 0,
        ii = i,
        jj = j,
        Pi = [],
        Pj = [];
      path[i][j] = [];
      if (i !== j) {
        while (ii != sourceBackboneIndex || P[ii] != sourceBackboneIndex) {
          Pi.push(ii);
          ii = P[ii];
        }
        while (jj != sourceBackboneIndex || P[jj] != sourceBackboneIndex) {
          Pj.push(jj);
          jj = P[jj];
        }
        let meet = sourceBackboneIndex;
        for (let m = 0; m < Pi.length; m++) {
          for (let n = 0; n < Pj.length; n++) {
            if (Pi[m] == Pj[n]) {
              meet = Pi[m];
              Pi.splice(m);
              Pj.splice(n);
            }
          }
        }
        path[i][j] = path[i][j].concat(Pi);
        path[i][j].push(meet);
        path[i][j] = path[i][j].concat(Pj.reverse());
      } else {
        path[i][j].push(i);
      }
      hop[i][j] = Pi.length + Pj.length;
    }
  }

 
  if (!MENTOR_1) {
    /* // Sử dụng lưu lượng gốc giữa các nút backbone
        P.forEach((e,i) => {
            if (e != i) {
                let n = Math.ceil(T_b[e][i]/C);
                let u = n == 0 ? 0 : T_b[e][i]/(n*C);
                links[e][i] = links[i][e] = {
                    n: n,
                    u: u,
                    d: calcDistance(backbones[e],backbones[i])
                }
                if (!HIDE_NO_TRAFFIC || T_b[e][i] > 0) {
                    lineNode(backbones[e], backbones[i], color(255, 204, 0));
                }
            }
        });
        */
    for (let i = 0; i < numberBackbones; i++) {
      for (let j = 0; j < numberBackbones; j++) {
        if (path[i][j].length > 1) {
          for (let k = 0; k < path[i][j].length - 1; k++) {
            linksTraffic[path[i][j][k]][path[i][j][k + 1]] += T_b[i][j];
            linksTraffic[path[i][j][k + 1]][path[i][j][k]] += T_b[i][j];
          }
        }
      }
    }

    P.forEach((e, i) => {
      if (e != i) {
        let trafic;
        if (MENTOR_2) {
          trafic = T_b[e][i];
        } else {
          trafic = linksTraffic[e][i];
        }
        let n = Math.ceil(trafic / C);
        let u = n == 0 ? 0 : trafic / (n * C);
        links[e][i] = links[i][e] = {
          n: n,
          u: u,
          d: calcDistance(backbones[e], backbones[i]),
          trafic: trafic,
        };
        if (!HIDE_NO_TRAFFIC || T_b[e][i] > 0) {
          if(MENTOR_2 == false){
            lineBlackboneNode(backbones[e], backbones[i], color(255, 204, 0));
          }
          
        }
      }
    });
  }

  //ve node
  drawCenterNode(points[min_moment_index], color(255, 69, 0));
  points.forEach((p, i) => {
    if (p.is_backbone && i != min_moment_index) drawBackboneNode(p);
  });

  // 3.3. Mentor 1: Chuyển lưu lượng
  if (MENTOR_1) {
    let pair = [];
    for (let i = 0; i < numberBackbones; i++) {
      for (let j = i; j < numberBackbones; j++) {
        pair.push({
          s: i,
          d: j,
          hop: hop[i][j],
        });
      }
    }
    pair.sort(function (a, b) {
      return b.hop - a.hop;
    });

    pair.forEach((e) => {
      let n = Math.ceil(T_b[e.s][e.d] / C);
      let u = T_b[e.s][e.d] / (n * C);
      if (e.hop >= 2 && T_b[e.s][e.d] > 0) {
        if (u >= u_min) {
          lineNode(backbones[e.s], backbones[e.d], color(111, 255, 0));
          links[e.s][e.d] = links[e.d][e.s] = {
            n: n,
            u: u,
            d: calcDistance(backbones[e.s], backbones[e.d]),
          };
        } else {
          let minDist = {
            dist: Infinity,
          };
          for (let i = 1; i < path[e.s][e.d].length - 1; i++) {
            let currentBackboneIndex = path[e.s][e.d][i];
            let localDist = calcDistance(
              backbones[e.s],
              backbones[currentBackboneIndex]
            );
            if (localDist < minDist.dist) {
              minDist = {
                dist: localDist,
                backboneIndex: currentBackboneIndex,
              };
            }
          }
          T_b[e.s][minDist.backboneIndex] = T_b[minDist.backboneIndex][e.s] =
            T_b[e.s][minDist.backboneIndex] + T_b[e.s][e.d];
          T_b[e.d][minDist.backboneIndex] = T_b[minDist.backboneIndex][e.d] =
            T_b[e.d][minDist.backboneIndex] + T_b[e.s][e.d];
        }
      } else if (e.hop == 1 && T_b[e.s][e.d] > 0) {
        lineNode(backbones[e.s], backbones[e.d], color(111, 255, 0));
        links[e.s][e.d] = links[e.d][e.s] = {
          n: n,
          u: u,
          d: calcDistance(backbones[e.s], backbones[e.d]),
        };
      }
    });
  }

  // 3.4. Mentor 2: Thêm liên kết trực tiếp
  if (MENTOR_2) {
    // 3.4.1. Tính khoảng cách giữa các nút backbone dọc theo cây
    let backbonesDistance = [];
    for (let i = 0; i < numberBackbones; i++) {
      backbonesDistance[i] = [];
      for (let j = 0; j < numberBackbones; j++) {
        backbonesDistance[i][j] = 0;
        for (let k = 0; k < path[i][j].length - 1; k++) {
          backbonesDistance[i][j] += calcDistance(
            backbones[path[i][j][k]],
            backbones[path[i][j][k + 1]]
          );
        }
      }
    }
    // Tìm các cặp nút có lưu lượng và số hop > 1
    let mentor2NodePair = [];
    for (let i = 0; i < numberBackbones - 1; i++) {
      for (let j = i + 1; j < numberBackbones; j++) {
        if (hop[i][j] > 1 && T_b[i][j] > 0) {
          let pair = [];
          pair.start = i;
          pair.stop = j;
          pair.numberHop = hop[i][j];
          pair.T = T_b[i][j];
          pair.distance = backbonesDistance[i][j];
          mentor2NodePair.push(pair);
        }
      }
    }
    // Sắp xếp các cặp nút theo thứ tự giảm dần về khoảng cách (theo cây)
    mentor2NodePair.sort(function (a, b) {
      return b.distance - a.distance;
    });
    // Phân loại s_list và d_list cho mỗi cặp nút
    mentor2NodePair.forEach((pair) => {
      let i = pair.start;
      let j = pair.stop;
      pair.sList = [];
      pair.dList = [];
      let L = calcDistance(backbones[i], backbones[j]);
      for (let k = 0; k < numberBackbones; k++) {
        let dist_ik = backbonesDistance[i][k];
        let dist_kj = backbonesDistance[k][j];
        if (dist_ik + L - dist_kj < 0) {
          pair.sList.push(k);
        }
        if (dist_kj + L < dist_ik) {
          pair.dList.push(k);
        }
      }
      pair.canAddDirectLink = true;
      let max = {
        value: 0,
        start: -1,
        stop: -1,
      };
      pair.sList.forEach((ni) => {
        pair.dList.forEach((nj) => {
          if (T_b[ni][nj] > 0 && (ni != i || nj != j)) {
            let localL =
              backbonesDistance[ni][nj] -
              backbonesDistance[ni][i] -
              backbonesDistance[j][nj];
            if (localL > max.value) {
              max.value = localL;
              max.start = ni;
              max.stop = nj;
            }
          }
        });
      });
      pair.maxL = max;
      if (max.value > calcDistance(backbones[i], backbones[j])) {
        if (max.value >= backbonesDistance[i][j]) {
          pair.canAddDirectLink = false;
        } else {
          if (Math.ceil(max.value) < backbonesDistance[i][j]) {
            backbonesDistance[i][j] = backbonesDistance[j][i] = Math.ceil(
              max.value
            );
          } else {
            backbonesDistance[i][j] = backbonesDistance[j][i] =
              (max.value + backbonesDistance[i][j]) / 2;
          }
        }
      } else {
        backbonesDistance[i][j] = backbonesDistance[j][i] = calcDistance(
          backbones[i],
          backbones[j]
        );
      }
    });
    // Vẽ các liên kết trực tiếp
    mentor2NodePair.forEach((pair) => {
      if (pair.canAddDirectLink) {
        lineNode(
          backbones[pair.start],
          backbones[pair.stop],
          color(255, 0, 111)
        );
        let n = Math.ceil(T_b[pair.start][pair.stop] / C);
        let u = T_b[pair.start][pair.stop] / (C * n);
        links[pair.start][pair.stop] = links[pair.stop][pair.start] = {
          n: n,
          u: u,
          d: backbonesDistance[pair.start][pair.stop],
          trafic: T_b[pair.start][pair.stop],
        };
      }
    });

    // In bảng liên kết trực tiếp
    let htmlDirectLinkTable =
      '<table class="ui celled table center aligned unstackable selectable"><thead><tr>';
    htmlDirectLinkTable +=
      "<th>Nút đấu</th><th>Nút cuối</th><th>Khoảng cách</th><th>maxL</th><th>Chiều dài cũ theo cây</th>";
    htmlDirectLinkTable +=
      "<th>Độ dài liên kết trực tiếp</th><th>Số đường</th><th>Giá</th>";
    htmlDirectLinkTable += "</tr></thead><tbody>";
    mentor2NodePair.forEach((pair) => {
      if (pair.canAddDirectLink) {
        htmlDirectLinkTable += "<tr><td>" + (pair.start + 1) + "</td>";
        htmlDirectLinkTable += "<td>" + (pair.stop + 1) + "</td>";
        htmlDirectLinkTable +=
          "<td>" +
          calcDistance(backbones[pair.start], backbones[pair.stop]).toFixed(2) +
          "</td>";
        htmlDirectLinkTable += "<td>" + pair.maxL.value.toFixed(2) + "</td>";
        htmlDirectLinkTable += "<td>" + pair.distance.toFixed(2) + "</td>";
        htmlDirectLinkTable +=
          "<td>" + links[pair.start][pair.stop].d.toFixed(2) + "</td>";
        htmlDirectLinkTable +=
          "<td>" + links[pair.start][pair.stop].n + "</td>";
        htmlDirectLinkTable +=
          "<td>" +
          (
            0.4 *
            links[pair.start][pair.stop].d *
            links[pair.start][pair.stop].n
          ).toFixed(2) +
          "</td></tr>";
      }
    });
    htmlDirectLinkTable += "</tbody></table>";
    $("#mentor2").show();
    $("#direct-link").html(htmlDirectLinkTable);
    $("#direct-link").scroll(function () {
      if ($(this).scrollTop() == 0) {
        $(this).css("border-top", "0");
      } else {
        $(this).css("border-top", "1px solid rgba(34,36,38,.1)");
      }
    });
  } else {
    $("#mentor2").hide();
  }

  let Price = 0;
  for (let i = 0; i < numberBackbones - 1; i++) {
    for (let j = i + 1; j < numberBackbones; j++) {
      Price += links[i][j].n * 0.4 * links[i][j].d;
    }
  }
  $("#price").html(Price.toFixed(2));
 

  // In danh sách liên kết
  let htmlLinkTable =
    '<table class="ui celled table center aligned unstackable selectable"><thead><tr>';
  htmlLinkTable +=
    "<th>Nút đầu</th><th>Nút cuối</th><th>Lưu lượng</th><th>Số đường</th>";
  htmlLinkTable += "<th>Độ sử dụng</th><th>Khoảng cách</th><th>Giá</th>";
  htmlLinkTable += "</tr></thead><tbody>";
  for (let i = 0; i < numberBackbones - 1; i++) {
    for (let j = i + 1; j < numberBackbones; j++) {
      if (links[i][j].n > 0) {
        htmlLinkTable += "<tr><td>" + (i + 1) + "</td>";
        htmlLinkTable += "<td>" + (j + 1) + "</td>";
        htmlLinkTable += "<td>" + links[i][j].trafic + "</td>";
        htmlLinkTable += "<td>" + links[i][j].n + "</td>";
        htmlLinkTable += "<td>" + links[i][j].u.toFixed(2) + "</td>";
        htmlLinkTable += "<td>" + links[i][j].d.toFixed(2) + "</td>";
        htmlLinkTable +=
          "<td>" +
          (0.4 * links[i][j].d * links[i][j].n).toFixed(2) +
          "</td></tr>";
      }
    }
  }
  htmlLinkTable +=
    '</tbody><tfoot><tr><th colspan="6">Tổng giá</th><th>' +
    Price.toFixed(2) +
    "</th></tr></tfoot></table>";
  $("#links").html(htmlLinkTable);
  $("#links").scroll(function () {
    if ($(this).scrollTop() == 0) {
      $(this).css("border-top", "0");
    } else {
      $(this).css("border-top", "1px solid rgba(34,36,38,.1)");
    }
  });

  let numberLine = [...Array(numberBackbones)].map((x) =>
    Array(numberBackbones).fill("-")
  );
  let usage = [...Array(numberBackbones)].map((x) =>
    Array(numberBackbones).fill("-")
  );
  for (let i = 0; i < numberBackbones; i++) {
    for (let j = i + 1; j < numberBackbones; j++) {
      if (links[i][j].n > 0) {
        numberLine[i][j] = numberLine[j][i] = links[i][j].n;
        usage[i][j] = usage[j][i] = links[i][j].u.toFixed(2);
      }
    }
  }
  $("#number-line").printMatrix({
    title: "n",
    matrix: numberLine,
    height: "300px",
  });
  $("#usage").printMatrix({
    title: "u",
    matrix: usage,
    height: "300px",
  });

  let htmlWeightsTable =
    '<table class="ui celled table center aligned unstackable selectable"><thead><tr><th>Nút</th><th>Trọng số w</th>';
  htmlWeightsTable += "</tr></thead><tbody>";
  points.forEach((element, index) => {
    htmlWeightsTable += "<tr><td>" + (index + 1) + "</td>";
    htmlWeightsTable +=
      "<td>" +
      element.w +
      "</td></tr>";
  });
  htmlWeightsTable += "</tbody></table>";
  $("#weight").html(htmlWeightsTable);

  points.forEach((e, i) => {
    if (e.is_access) {
      fill(color(206, 132, 98));
      textSize(13);
      let x = Math.round(e.x / 2) - 10;
      let y = Math.round(e.y / 2) - 5;
      x = x < 10 ? 5 : x > x_matrix / 2 - 20 ? x_matrix / 2 - 20 : x;
      y = y < 10 ? y + 15 : y;
      stroke(0);
      text(i + 1, x, y);
    }
  });
  backbones.forEach((e, i) => {
    fill(color(209, 255, 82));
    textSize(15);
    let x = Math.round(e.x / 2) - 20;
    let y = Math.round(e.y / 2) - 20;
    x = x < 10 ? 5 : x > x_matrix / 2 - 40 ? x_matrix / 2 - 40 : x;
    y = y < 30 ? y + 45 : y;
    stroke(0);
    text(i + 1 + "(" + (e.index + 1) + ")", x, y);
  });

  $(".ui.sticky").sticky({
    offset: 50,
  });


  let newPoints = points.map(point => {
    if(point.is_access){
      let backboneNode = {
        // Xử lý từng phần tử trong mảng 'backbone'
        // Ví dụ: Trích xuất các thuộc tính bạn muốn từ 'accessItem'
        
        x: point.backbone.x,
        y: point.backbone.y,
        index: point.backbone.index,
        w: point.backbone.w,
        isAccess: point.backbone.is_access,
        backboneDistance: point.backbone.backbone_distance,
        moment: point.backbone.moment,
       
      };
    
      // Tạo đối tượng mới chứa các thuộc tính cần thiết và chuỗi JSON của 'access'
      return {
        x: point.x,
        y: point.y,
        index: point.index,
        w: point.w,
        isAccess: point.is_access,
        isBackbone: point.is_backbone,
        backboneDistance: point.backbone_distance,
        moment: point.moment,
        backbone: backboneNode // Chuyển đổi 'processedAccess' thành chuỗi JSON
      };
      
      
    }
    else{
      return {
        x: point.x,
        y: point.y,
        index: point.index,
        w: point.w,
        isAccess: point.is_access,
        isBackbone: point.is_backbone,
        backboneDistance: point.backbone_distance,
        moment: point.moment,
        
      };
    }
    
    // Xử lý mảng 'backbone' trong 'point.access' trước khi chuyển đổi thành chuỗi JSON
   
  });
  let data = {
    myPoints : newPoints ,
    
    tb: [...T_b]
  }
  function flattenObject(obj, parentKey = '', result = {}) {
    for (let key in obj) {
      let newKey = parentKey ? `${parentKey}.${key}` : key;
      if (typeof obj[key] === 'object' && obj[key] !== null && !Array.isArray(obj[key])) {
        flattenObject(obj[key], newKey, result);
      } else {
        result[newKey] = obj[key];
      }
    }
    return result;
  }
 
  fetch('https://localhost:7163/createfile', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json'
  },
  body: JSON.stringify(data)
})
.then(response => {
  return response.blob(); // Chuyển đổi dữ liệu nhận được thành Blob
})
.then(blob => {
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'Sample.xlsx'; // Tên file
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  window.URL.revokeObjectURL(url);
})
.catch(error => {
  console.error('Đã xảy ra lỗi:', error);
});
}
