/**
 * 🎯 PURE COORDINATE RENDERER WITH PERSISTENT CHART SHEETS
 * Charts are created in separate sheets and kept for reference
 */

const CONFIG = {
  PRESENTATION_NAME: 'Business Review',
  //SLIDE_TEMPLATES:['Slide1', 'Slide2', 'Slide3', 'Slide4', 'Slide5', 'Slide6', 'Slide7', 'Slide8', 'Slide9', 'Slide10', 'Slide11', 'Slide12', 'Slide13', 'Slide14', 'Slide15', 'Slide16', 'Slide17', 'Slide17A', 'Slide18' ] 

  //SLIDE_TEMPLATES:['Slide17A_QuarterlySuccessPlanning']
  //SLIDE_TEMPLATES:['Slide1_overview','Slide2_meetingGoals','Slide3_Agenda','Slide4_DoceboTeam','Slide5_ClientTeam','Slide6_ExecutiveSummary','Slide7_CurrentStateLearning','Slide8_LearningSolutionOverview','Slide9_PartnershipTimeline','Slide10_KeyStrategic','Slide11_UsageAndMetrics','Slide12_YourDoceboMetrics','Slide13_LearnerEngagement','Slide14_LearningSession', 'Slide15_ContentEffectiveness','Slide16_OptimizationAndSuccessPlan','Slide17_OptimizationStrategyTable','Slide17A_QuarterlySuccessPlanning','Slide18_Grazie','Slide19_Appendix'] 

  SLIDE_TEMPLATES:['Slide4_DoceboTeam','Slide5_ClientTeam'] 

  //SLIDE_TEMPLATES:['Slide17A_QuarterlySuccessPlanning']
  //SLIDE_TEMPLATES:['check']
};

const SHEETS_CONFIG = {
  SPREADSHEET_ID: '10q6kFhPycyfF_Kau6d0kmXzhdqhiGMisvTd89byBmSI', // Replace with your actual spreadsheet ID
  SHEET_NAME: 'Chart Data',
  CHART_SHEET_NAME: 'Charts'
};

const PLACEHOLDER_IMAGE_URL = 'https://qa-uploads.unifyapps.com/user-uploads/1/1770225007615/Screenshot_2026-02-04_at_10.39.xn--46PM-1d7a.png';

// Global variable to track chart count
let chartCounter = 0;

function createCoordinateSlides() {

  let parameters = {
      "accountName": "Aetrex Worldwide, Inc.",
      "accountTeamArray": [
        {
          "area": "Enter Area Here",
          "impact": "Enter Impact Here",
          "membership": "Enter Role Here",
          "name": "Enter Name Here",
          "stage": "Enter Stage Here",
          "useCase": "Enter Use Case Here"
        }
      ],
      "ceChart": "https://s3.ap-south-1.amazonaws.com/uat.unifyapps.com.cloudstorage.722/workflow_uploads/722/image15.png?response-content-disposition=attachment%3Bfilename%3D%22image15.png%22&X-Amz-Security-Token=IQoJb3JpZ2luX2VjEEgaCmFwLXNvdXRoLTEiRjBEAiBWKmUcXHmKTQzwsmzoTvYO6BJCmm%2BPaNBbUEwya4Vd%2BAIgMj9AoF%2FheVEXNJfnru4XVNuUU3yJO5Z41Ujz2w2KQrAqvwUIERAAGgw1NDY5MTgxMDYyODMiDOwh2B4AO8t%2FoO1LeSqcBb9k8ksXZc7%2BMF7PUJgeEz1Ew7pdyYGeUBo1eCW71jnUBsmU%2FpscTNUWr7588q7ZswgsGINMCpbB9EiSFuNEUa5vzaVxK7pyNXXcvUqkg5FowlQzTtWWyaQukejw3Ku2J2%2BybiVKQHR8DAq982HfG7N0vG6SjE7%2Fmg1HTTU%2FleHEow%2BLD7HreTVe2wJIQnIKMaSHLKstLtwAQ%2F8KII2dz75g05pzJszDZyZnMxEVX1ZVeE0u85LRDdvvK3Ga5Bw5R%2BVBabPzC%2BgXEvhq7gj9tP5Xav9jnwgPSy1vCOQgcu2QkEK0K6DB6M56MqMsF4jxQkPEsNEKuPxDaWYZqkgPdjeUawoLWltMqrE8gLs1iIY2ZkYmkpyKKwfumQOPjg4Dqqr5F26JYz7fYgFS1b7SwN0jdDMYzSYlZnuaOFTReKE7TVkipwqE8DHLrCPeBsYkYSF2I6rC4GALIYq%2FquNg3Iqo8TbtkVaYwRE98JkmzOcnbVUOcZDvD96MAVQBrxPAN%2BuJT%2BctCibXd1tA8SzQAVUOPWAzx0U%2BLgC5fX3e61QUXT8xdRiCsSGcUD1gexaCmn3UQ0O6FKz7BC9N8fsOmdmFag4kKNo13ETaHqsy1bZFNMi%2FYmjObPDY0ogQl01k4nEY0pYM%2BgGfjRIGdK9Yn5iIns9d3MHN9g9weyLVL7rZ6EhQwMY00hx%2FA5W1CsIo%2BJzCkVc9OF%2B%2BMVBbjEFxMIgSKGW1OklFPbyFvsw2%2F6vN77Zt0U1CJcVLvf5xqPqWY0jY%2BY3Xd8VdRtRU9xhH%2FxOV1r%2BROjBMqkgHBpH0kRzE6jagBZVWHv46LEBS6bJp5H3xBsv9flfvxeUKHTAi10cUwpMXmrhQBz%2BERaNZqsK4kPFBOp7zpUJ%2BsWedMJvvi8wGOrIBToUA9hku3I0LSwWSq3niRC0x8A8oNz0v5xsSVijM2Xs29eGiZX0OdLmz%2BVSU0dp5fxI9pCaIhF2GSoJZHGJxU6J2kwkcYxeNP4IyYuKkQODOyEzBSlLnRtnivBCuoK0x0%2BsujkmOrRjUVfagSxaXF2UZ9fH2ZAU3t%2Bx8xhM4OfrmxmDYT7QQiC%2BZAk0YjI0ZIxB6AtQBVYQZ%2BZh451sQXcSzIut%2FStxRL%2B0VgNgiJoSTmw%3D%3D&X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Date=20260204T074618Z&X-Amz-SignedHeaders=host&X-Amz-Credential=ASIAX6VW4ASV4E5YK7WW%2F20260204%2Fap-south-1%2Fs3%2Faws4_request&X-Amz-Expires=86400&X-Amz-Signature=8adf0b9103a934bcd34b72e5f8fa79c833a097d130b74453efff2babe83596a7",
      "ceContent": "<ul><li><p><span style=\"font-size: 20px\">E-Learning is the only enrollment method utilized throughout the year.</span></p></li><li><p><span style=\"font-size: 20px\">June 2025 saw the highest e-learning enrollments at 1,056.</span></p></li><li><p><span style=\"font-size: 20px\">Significant drop in enrollments from June to July 2025.</span></p></li><li><p><span style=\"font-size: 20px\">No activity recorded for ILT, Self, Internal, Admin, or Docebo enrollments.</span></p></li><li><p><span style=\"font-size: 20px\">Recent months (November–December 2025) indicate declining engagement.</span></p></li></ul>",
      "date": "3 Dec 2025",
      "date1": "03-2025",
      "date2": "06-2025",
      "date3": "09-2025",
      "date4": "12-2025",
      "date5": "03-2026",
      "doceboTeamArray": [
        {
          "Citations": "[2,3,4]",
          "membership": "Account Manager",
          "name": "Stephanie Sloss",
          "profilePicture": "Unavailable"
        },
        {
          "Citations": "[2,3,4]",
          "membership": "Customer Success Manager",
          "name": "CSM NA POOL",
          "profilePicture": "Unavailable"
        }
      ],
      "goal1": "Enhance sales enablement training effectiveness and engagement",
      "goal2": "Optimize onboarding and professional development programs",
      "goal3": "Expand customer education and certification programs",
      "goal4": "Strengthen partner training and enablement initiatives",
      "goal5": "Implement advanced learning analytics and reporting",
      "headerUrl": "Enter Your Text Here",
      "imageUrl": "https://t4.ftcdn.net/jpg/03/32/41/57/360_F_332415747_wgPEZYABPCoBC2MbbZMsK9NHEguootsb.jpg",
      "logoUrl": "https://zenprospect-production.s3.amazonaws.com/uploads/pictures/6858c73b879649000191607d/picture",
      "lookingValue": "Planning enhanced customer education initiatives and expanded partner training programs",
      "lsChart": "https://s3.ap-south-1.amazonaws.com/uat.unifyapps.com.cloudstorage.722/workflow_uploads/722/image14.png?response-content-disposition=attachment%3Bfilename%3D%22image14.png%22&X-Amz-Security-Token=IQoJb3JpZ2luX2VjEEgaCmFwLXNvdXRoLTEiRjBEAiBWKmUcXHmKTQzwsmzoTvYO6BJCmm%2BPaNBbUEwya4Vd%2BAIgMj9AoF%2FheVEXNJfnru4XVNuUU3yJO5Z41Ujz2w2KQrAqvwUIERAAGgw1NDY5MTgxMDYyODMiDOwh2B4AO8t%2FoO1LeSqcBb9k8ksXZc7%2BMF7PUJgeEz1Ew7pdyYGeUBo1eCW71jnUBsmU%2FpscTNUWr7588q7ZswgsGINMCpbB9EiSFuNEUa5vzaVxK7pyNXXcvUqkg5FowlQzTtWWyaQukejw3Ku2J2%2BybiVKQHR8DAq982HfG7N0vG6SjE7%2Fmg1HTTU%2FleHEow%2BLD7HreTVe2wJIQnIKMaSHLKstLtwAQ%2F8KII2dz75g05pzJszDZyZnMxEVX1ZVeE0u85LRDdvvK3Ga5Bw5R%2BVBabPzC%2BgXEvhq7gj9tP5Xav9jnwgPSy1vCOQgcu2QkEK0K6DB6M56MqMsF4jxQkPEsNEKuPxDaWYZqkgPdjeUawoLWltMqrE8gLs1iIY2ZkYmkpyKKwfumQOPjg4Dqqr5F26JYz7fYgFS1b7SwN0jdDMYzSYlZnuaOFTReKE7TVkipwqE8DHLrCPeBsYkYSF2I6rC4GALIYq%2FquNg3Iqo8TbtkVaYwRE98JkmzOcnbVUOcZDvD96MAVQBrxPAN%2BuJT%2BctCibXd1tA8SzQAVUOPWAzx0U%2BLgC5fX3e61QUXT8xdRiCsSGcUD1gexaCmn3UQ0O6FKz7BC9N8fsOmdmFag4kKNo13ETaHqsy1bZFNMi%2FYmjObPDY0ogQl01k4nEY0pYM%2BgGfjRIGdK9Yn5iIns9d3MHN9g9weyLVL7rZ6EhQwMY00hx%2FA5W1CsIo%2BJzCkVc9OF%2B%2BMVBbjEFxMIgSKGW1OklFPbyFvsw2%2F6vN77Zt0U1CJcVLvf5xqPqWY0jY%2BY3Xd8VdRtRU9xhH%2FxOV1r%2BROjBMqkgHBpH0kRzE6jagBZVWHv46LEBS6bJp5H3xBsv9flfvxeUKHTAi10cUwpMXmrhQBz%2BERaNZqsK4kPFBOp7zpUJ%2BsWedMJvvi8wGOrIBToUA9hku3I0LSwWSq3niRC0x8A8oNz0v5xsSVijM2Xs29eGiZX0OdLmz%2BVSU0dp5fxI9pCaIhF2GSoJZHGJxU6J2kwkcYxeNP4IyYuKkQODOyEzBSlLnRtnivBCuoK0x0%2BsujkmOrRjUVfagSxaXF2UZ9fH2ZAU3t%2Bx8xhM4OfrmxmDYT7QQiC%2BZAk0YjI0ZIxB6AtQBVYQZ%2BZh451sQXcSzIut%2FStxRL%2B0VgNgiJoSTmw%3D%3D&X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Date=20260204T074618Z&X-Amz-SignedHeaders=host&X-Amz-Credential=ASIAX6VW4ASV4E5YK7WW%2F20260204%2Fap-south-1%2Fs3%2Faws4_request&X-Amz-Expires=86400&X-Amz-Signature=53e83d1f9c428301af2bb3586d81a98e7db0717715056bd1e417df4526d9edf9",
      "lsContent": "<ul><li><p><span style=\"font-size: 20px\">All learning time is from E Learning, with ILT time at zero throughout.</span></p></li><li><p><span style=\"font-size: 20px\">Peak learning time was in March 2025 at 44,160 minutes, then dropped in April 2025.</span></p></li><li><p><span style=\"font-size: 20px\">Consistent engagement above 30,000 minutes from May to December 2025, except in April 2025.</span></p></li><li><p><span style=\"font-size: 20px\">No learning time recorded for January 2026; trend declines after December 2025.</span></p></li></ul>",
      "mauChart": "https://s3.ap-south-1.amazonaws.com/uat.unifyapps.com.cloudstorage.722/workflow_uploads/722/image13.png?response-content-disposition=attachment%3Bfilename%3D%22image13.png%22&X-Amz-Security-Token=IQoJb3JpZ2luX2VjEEgaCmFwLXNvdXRoLTEiRjBEAiBWKmUcXHmKTQzwsmzoTvYO6BJCmm%2BPaNBbUEwya4Vd%2BAIgMj9AoF%2FheVEXNJfnru4XVNuUU3yJO5Z41Ujz2w2KQrAqvwUIERAAGgw1NDY5MTgxMDYyODMiDOwh2B4AO8t%2FoO1LeSqcBb9k8ksXZc7%2BMF7PUJgeEz1Ew7pdyYGeUBo1eCW71jnUBsmU%2FpscTNUWr7588q7ZswgsGINMCpbB9EiSFuNEUa5vzaVxK7pyNXXcvUqkg5FowlQzTtWWyaQukejw3Ku2J2%2BybiVKQHR8DAq982HfG7N0vG6SjE7%2Fmg1HTTU%2FleHEow%2BLD7HreTVe2wJIQnIKMaSHLKstLtwAQ%2F8KII2dz75g05pzJszDZyZnMxEVX1ZVeE0u85LRDdvvK3Ga5Bw5R%2BVBabPzC%2BgXEvhq7gj9tP5Xav9jnwgPSy1vCOQgcu2QkEK0K6DB6M56MqMsF4jxQkPEsNEKuPxDaWYZqkgPdjeUawoLWltMqrE8gLs1iIY2ZkYmkpyKKwfumQOPjg4Dqqr5F26JYz7fYgFS1b7SwN0jdDMYzSYlZnuaOFTReKE7TVkipwqE8DHLrCPeBsYkYSF2I6rC4GALIYq%2FquNg3Iqo8TbtkVaYwRE98JkmzOcnbVUOcZDvD96MAVQBrxPAN%2BuJT%2BctCibXd1tA8SzQAVUOPWAzx0U%2BLgC5fX3e61QUXT8xdRiCsSGcUD1gexaCmn3UQ0O6FKz7BC9N8fsOmdmFag4kKNo13ETaHqsy1bZFNMi%2FYmjObPDY0ogQl01k4nEY0pYM%2BgGfjRIGdK9Yn5iIns9d3MHN9g9weyLVL7rZ6EhQwMY00hx%2FA5W1CsIo%2BJzCkVc9OF%2B%2BMVBbjEFxMIgSKGW1OklFPbyFvsw2%2F6vN77Zt0U1CJcVLvf5xqPqWY0jY%2BY3Xd8VdRtRU9xhH%2FxOV1r%2BROjBMqkgHBpH0kRzE6jagBZVWHv46LEBS6bJp5H3xBsv9flfvxeUKHTAi10cUwpMXmrhQBz%2BERaNZqsK4kPFBOp7zpUJ%2BsWedMJvvi8wGOrIBToUA9hku3I0LSwWSq3niRC0x8A8oNz0v5xsSVijM2Xs29eGiZX0OdLmz%2BVSU0dp5fxI9pCaIhF2GSoJZHGJxU6J2kwkcYxeNP4IyYuKkQODOyEzBSlLnRtnivBCuoK0x0%2BsujkmOrRjUVfagSxaXF2UZ9fH2ZAU3t%2Bx8xhM4OfrmxmDYT7QQiC%2BZAk0YjI0ZIxB6AtQBVYQZ%2BZh451sQXcSzIut%2FStxRL%2B0VgNgiJoSTmw%3D%3D&X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Date=20260204T074618Z&X-Amz-SignedHeaders=host&X-Amz-Credential=ASIAX6VW4ASV4E5YK7WW%2F20260204%2Fap-south-1%2Fs3%2Faws4_request&X-Amz-Expires=86400&X-Amz-Signature=80959c3816aedf8b5d1449469fc141f8a46a80efaccc8dcc39da04368b1dfc80",
      "mauContent": "<ul><li><p><span style=\"font-size: 20px\">April and June 2025 exceeded the plan target of 300 users.</span></p></li><li><p><span style=\"font-size: 20px\">Significant drop in active users noted from December 2025 to January 2026.</span></p></li><li><p><span style=\"font-size: 20px\">Overall YoY change shows a reduction of 56.7% in active users.</span></p></li></ul>",
      "metrics1": "Training completion rates and sales performance correlation",
      "metrics2": "Time-to-productivity and employee satisfaction scores",
      "metrics3": "Customer engagement rates and certification completions",
      "metrics4": "Partner activation rates and performance metrics",
      "metrics5": "Learning ROI measurement and business impact tracking",
      "meetGoalOne": "<span style=\"font-size: 14px\">· Align on your strategic goals and vision in partnership with Docebo.</span>",
      "meetGoalTwo": "<span style=\"font-size: 14px\">· Showcase value in the tools you have access to, provide optimization recommendations and relevant product updates.</span>",
      "meetGoalThree": "<span style=\"font-size: 14px\">· Success planning.</span>",
      "key1": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "key2": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "key3": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "key4": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "key5": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "custOwner1": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "custOwner2": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "custOwner3": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "custOwner4": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "custOwner5": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "doceboOwner1": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "doceboOwner2": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "doceboOwner3": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "doceboOwner4": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "doceboOwner5": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "priority1": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "priority2": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "priority3": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "priority4": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "priority5": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "status1": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "status2": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "status3": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "status4": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "status5": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "dueDate1": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "dueDate2": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "dueDate3": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "dueDate4": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "dueDate5": "<span style=\"font-size: 9px\">Enter you text here</span>",
      "milestone1": "Platform optimization initiatives",
      "milestone2": "Customer education expansion",
      "milestone3": "Partner training enhancement",
      "milestone4": "Advanced analytics implementation",
      "milestone5": "Learning ecosystem integration",
      "outcome1": "Engagement",
      "outcome10": "Effectiveness",
      "outcome2": "Retention",
      "outcome3": "Performance",
      "outcome4": "Productivity",
      "outcome5": "Satisfaction",
      "outcome6": "Certification",
      "outcome7": "Activation",
      "outcome8": "ROI",
      "outcome9": "Impact",
      "productValue": "Leveraging comprehensive learning platform for sales enablement and professional development programs",
      "progressImg1": "https://s3.ap-south-1.amazonaws.com/uat.unifyapps.com.cloudstorage.722/workflow_uploads/722/image17.png?response-content-disposition=attachment%3Bfilename%3D%22image17.png%22&X-Amz-Security-Token=IQoJb3JpZ2luX2VjEEgaCmFwLXNvdXRoLTEiRjBEAiBWKmUcXHmKTQzwsmzoTvYO6BJCmm%2BPaNBbUEwya4Vd%2BAIgMj9AoF%2FheVEXNJfnru4XVNuUU3yJO5Z41Ujz2w2KQrAqvwUIERAAGgw1NDY5MTgxMDYyODMiDOwh2B4AO8t%2FoO1LeSqcBb9k8ksXZc7%2BMF7PUJgeEz1Ew7pdyYGeUBo1eCW71jnUBsmU%2FpscTNUWr7588q7ZswgsGINMCpbB9EiSFuNEUa5vzaVxK7pyNXXcvUqkg5FowlQzTtWWyaQukejw3Ku2J2%2BybiVKQHR8DAq982HfG7N0vG6SjE7%2Fmg1HTTU%2FleHEow%2BLD7HreTVe2wJIQnIKMaSHLKstLtwAQ%2F8KII2dz75g05pzJszDZyZnMxEVX1ZVeE0u85LRDdvvK3Ga5Bw5R%2BVBabPzC%2BgXEvhq7gj9tP5Xav9jnwgPSy1vCOQgcu2QkEK0K6DB6M56MqMsF4jxQkPEsNEKuPxDaWYZqkgPdjeUawoLWltMqrE8gLs1iIY2ZkYmkpyKKwfumQOPjg4Dqqr5F26JYz7fYgFS1b7SwN0jdDMYzSYlZnuaOFTReKE7TVkipwqE8DHLrCPeBsYkYSF2I6rC4GALIYq%2FquNg3Iqo8TbtkVaYwRE98JkmzOcnbVUOcZDvD96MAVQBrxPAN%2BuJT%2BctCibXd1tA8SzQAVUOPWAzx0U%2BLgC5fX3e61QUXT8xdRiCsSGcUD1gexaCmn3UQ0O6FKz7BC9N8fsOmdmFag4kKNo13ETaHqsy1bZFNMi%2FYmjObPDY0ogQl01k4nEY0pYM%2BgGfjRIGdK9Yn5iIns9d3MHN9g9weyLVL7rZ6EhQwMY00hx%2FA5W1CsIo%2BJzCkVc9OF%2B%2BMVBbjEFxMIgSKGW1OklFPbyFvsw2%2F6vN77Zt0U1CJcVLvf5xqPqWY0jY%2BY3Xd8VdRtRU9xhH%2FxOV1r%2BROjBMqkgHBpH0kRzE6jagBZVWHv46LEBS6bJp5H3xBsv9flfvxeUKHTAi10cUwpMXmrhQBz%2BERaNZqsK4kPFBOp7zpUJ%2BsWedMJvvi8wGOrIBToUA9hku3I0LSwWSq3niRC0x8A8oNz0v5xsSVijM2Xs29eGiZX0OdLmz%2BVSU0dp5fxI9pCaIhF2GSoJZHGJxU6J2kwkcYxeNP4IyYuKkQODOyEzBSlLnRtnivBCuoK0x0%2BsujkmOrRjUVfagSxaXF2UZ9fH2ZAU3t%2Bx8xhM4OfrmxmDYT7QQiC%2BZAk0YjI0ZIxB6AtQBVYQZ%2BZh451sQXcSzIut%2FStxRL%2B0VgNgiJoSTmw%3D%3D&X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Date=20260204T074618Z&X-Amz-SignedHeaders=host&X-Amz-Credential=ASIAX6VW4ASV4E5YK7WW%2F20260204%2Fap-south-1%2Fs3%2Faws4_request&X-Amz-Expires=86400&X-Amz-Signature=8700a3c6c6ece198172dfe85ba38d863ce922404362e3d2065f0bc5ec2b4486f",
      "progressImg3": "Enter Your Text Here",
      "progressImg4": "Enter Your Text Here",
      "renewalDate": "Enter Your Text Here",
      "slide10Img": "https://s3.ap-south-1.amazonaws.com/uat.unifyapps.com.cloudstorage.722/workflow_uploads/722/image9.png?response-content-disposition=attachment%3Bfilename%3D%22image9.png%22&X-Amz-Security-Token=IQoJb3JpZ2luX2VjEEgaCmFwLXNvdXRoLTEiRjBEAiBWKmUcXHmKTQzwsmzoTvYO6BJCmm%2BPaNBbUEwya4Vd%2BAIgMj9AoF%2FheVEXNJfnru4XVNuUU3yJO5Z41Ujz2w2KQrAqvwUIERAAGgw1NDY5MTgxMDYyODMiDOwh2B4AO8t%2FoO1LeSqcBb9k8ksXZc7%2BMF7PUJgeEz1Ew7pdyYGeUBo1eCW71jnUBsmU%2FpscTNUWr7588q7ZswgsGINMCpbB9EiSFuNEUa5vzaVxK7pyNXXcvUqkg5FowlQzTtWWyaQukejw3Ku2J2%2BybiVKQHR8DAq982HfG7N0vG6SjE7%2Fmg1HTTU%2FleHEow%2BLD7HreTVe2wJIQnIKMaSHLKstLtwAQ%2F8KII2dz75g05pzJszDZyZnMxEVX1ZVeE0u85LRDdvvK3Ga5Bw5R%2BVBabPzC%2BgXEvhq7gj9tP5Xav9jnwgPSy1vCOQgcu2QkEK0K6DB6M56MqMsF4jxQkPEsNEKuPxDaWYZqkgPdjeUawoLWltMqrE8gLs1iIY2ZkYmkpyKKwfumQOPjg4Dqqr5F26JYz7fYgFS1b7SwN0jdDMYzSYlZnuaOFTReKE7TVkipwqE8DHLrCPeBsYkYSF2I6rC4GALIYq%2FquNg3Iqo8TbtkVaYwRE98JkmzOcnbVUOcZDvD96MAVQBrxPAN%2BuJT%2BctCibXd1tA8SzQAVUOPWAzx0U%2BLgC5fX3e61QUXT8xdRiCsSGcUD1gexaCmn3UQ0O6FKz7BC9N8fsOmdmFag4kKNo13ETaHqsy1bZFNMi%2FYmjObPDY0ogQl01k4nEY0pYM%2BgGfjRIGdK9Yn5iIns9d3MHN9g9weyLVL7rZ6EhQwMY00hx%2FA5W1CsIo%2BJzCkVc9OF%2B%2BMVBbjEFxMIgSKGW1OklFPbyFvsw2%2F6vN77Zt0U1CJcVLvf5xqPqWY0jY%2BY3Xd8VdRtRU9xhH%2FxOV1r%2BROjBMqkgHBpH0kRzE6jagBZVWHv46LEBS6bJp5H3xBsv9flfvxeUKHTAi10cUwpMXmrhQBz%2BERaNZqsK4kPFBOp7zpUJ%2BsWedMJvvi8wGOrIBToUA9hku3I0LSwWSq3niRC0x8A8oNz0v5xsSVijM2Xs29eGiZX0OdLmz%2BVSU0dp5fxI9pCaIhF2GSoJZHGJxU6J2kwkcYxeNP4IyYuKkQODOyEzBSlLnRtnivBCuoK0x0%2BsujkmOrRjUVfagSxaXF2UZ9fH2ZAU3t%2Bx8xhM4OfrmxmDYT7QQiC%2BZAk0YjI0ZIxB6AtQBVYQZ%2BZh451sQXcSzIut%2FStxRL%2B0VgNgiJoSTmw%3D%3D&X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Date=20260204T074617Z&X-Amz-SignedHeaders=host&X-Amz-Credential=ASIAX6VW4ASV4E5YK7WW%2F20260204%2Fap-south-1%2Fs3%2Faws4_request&X-Amz-Expires=86400&X-Amz-Signature=1096c404d59c308def1d5dd9427d7e02e7d7c9b514a30055edf0f8ec6765f0b9",
      "solutions1": "Learn",
      "solutions2": "Connect",
      "solutions3": "Academy",
      "strategicValue": "Driving comprehensive learning transformation across sales enablement and customer education initiatives",
      "tableData": [
        {
          "actionPlan": "Deploy advanced sales training modules and analytics",
          "area": "Sales Enablement Enhancement",
          "features": "Advanced learning paths and performance tracking",
          "impact": "Improved sales performance and faster onboarding",
          "stage": "Implementing",
          "useCase": "Internal sales team training and performance improvement programs"
        },
        {
          "actionPlan": "Launch comprehensive customer education portal",
          "area": "Customer Education Expansion",
          "features": "Extended enterprise and certification management",
          "impact": "Enhanced customer satisfaction and product adoption",
          "stage": "Adopting",
          "useCase": "External customer training and certification programs"
        },
        {
          "actionPlan": "Implement comprehensive learning analytics dashboard",
          "area": "Learning Analytics Implementation",
          "features": "Advanced reporting and learning analytics",
          "impact": "Data-driven learning decisions and ROI measurement",
          "stage": "Not started",
          "useCase": "Internal performance measurement and ROI tracking"
        },
        {
          "actionPlan": "Enhance partner training programs and tracking",
          "area": "Partner Training Optimization",
          "features": "Partner portal and certification tracking",
          "impact": "Improved partner performance and engagement",
          "stage": "Adopting",
          "useCase": "External partner enablement and certification programs"
        }
      ],
      "techStack1": "Enter Tool 1",
      "techStack2": "Enter Tool 2",
      "techStack3": "Enter Tool 3",
      "useCase1": "Sales",
      "useCase2": "Onboarding",
      "useCase3": "Customer",
      "winsValue": "Successfully established multi-faceted learning platform supporting sales, onboarding and customer education"
    }
      
  
  try {
    Logger.log('🚀 Creating pure coordinate slides...');

    if (!parameters || typeof parameters !== 'object') {
      Logger.log('❌ Invalid parameters provided');
      return { success: false, error: 'Invalid parameters provided' };
    }
    
    // Reset chart counter
    chartCounter = 0;
    
    const presentation = SlidesApp.create(CONFIG.PRESENTATION_NAME);
    Logger.log(`✅ Presentation created: ${presentation.getName()}`);
    
    const defaultSlide = presentation.getSlides()[0];
    defaultSlide.remove();
    
    const actualWidth = presentation.getPageWidth();
    const actualHeight = presentation.getPageHeight();
    
    Logger.log(`📐 Google Slides dimensions: ${actualWidth} x ${actualHeight}`);
    
    CONFIG.SLIDE_TEMPLATES.forEach((templateName, index) => {
      try {
        Logger.log(`🎨 Creating slide ${index + 1} from ${templateName} coordinates...`);
        
        const slide = presentation.appendSlide();
        
        if (!slide) {
          Logger.log(`❌ CRITICAL: Failed to create slide!`);
          return;
        }
        
        Logger.log(`✅ Slide created successfully, ObjectId: ${slide.getObjectId()}`);
        
        const htmlContent = HtmlService.createTemplateFromFile(templateName).evaluate().getContent();
        
        const coordinateData = extractCoordinateData(htmlContent, parameters);
        if (coordinateData) {
          renderCoordinateElements(slide, coordinateData, templateName, presentation);
          Logger.log(`✅ Slide ${index + 1} created with ${coordinateData.elements.length} elements`);
        } else {
          Logger.log(`⚠️ No coordinate data found in ${templateName}`);
        }
        
      } catch (slideError) {
        Logger.log(`❌ Error creating slide ${index + 1}: ${slideError.toString()}`);
        Logger.log(`Stack trace: ${slideError.stack}`);
      }
    });
    
    sharePresentationWithSpecificEmails(presentation)
    Logger.log(`✅ All coordinate slides created`);
    Logger.log(`📊 Total charts created: ${chartCounter}`);
    Logger.log(`🔗 URL: ${presentation.getUrl()}`);
    
    return {
      success: true,
      url: presentation.getUrl(),
      slidesCreated: CONFIG.SLIDE_TEMPLATES.length,
      chartsCreated: chartCounter,
    };
    
  } catch (error) {
    Logger.log(`❌ Error: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
    return { success: false, error: error.toString() };
  }
}

function cleanTextForJson(str) {
  if (!str) return '';
  
  Logger.log(`🧹 Cleaning HTML text: "${String(str).substring(0, 100)}..."`);
  
  // OPTION 1: Strip HTML tags completely (recommended for text display)
  let cleaned = stripHtmlTags(str);
  
  // OPTION 2: If you need to preserve HTML, use escapeHtmlForJson instead
  // let cleaned = escapeHtmlForJson(str);
  
  Logger.log(`✅ Cleaned result: "${cleaned}"`);
  return cleaned;
}

function stripHtmlTags(html) {
  if (!html) return '';
  
  // Convert any data type to string safely
  let htmlString = '';
  if (typeof html === 'string') {
    htmlString = html;
  } else if (typeof html === 'object') {
    htmlString = JSON.stringify(html);  // Convert objects to JSON
  } else {
    htmlString = String(html);          // Convert other types to string
  }
  
  // Now safely call .replace()
  let text = htmlString.replace(/<[^>]*>/g, '');
  
  // Decode HTML entities
  text = text
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&nbsp;/g, ' ')
    .replace(/&hellip;/g, '...')
    .replace(/&ndash;/g, '-')
    .replace(/&mdash;/g, '--')
    .replace(/&rsquo;/g, "'")
    .replace(/&lsquo;/g, "'")
    .replace(/&rdquo;/g, '"')
    .replace(/&ldquo;/g, '"');
  
  // Clean up whitespace and control characters
  text = text
    .replace(/[\x00-\x1F\x7F-\x9F]/g, '') // Remove control characters
    .replace(/\s+/g, ' ') // Normalize whitespace
    .trim();
  
  // Escape for JSON
  text = text
    .replace(/\\/g, '\\\\')    // Escape backslashes
    .replace(/"/g, '\\"')      // Escape quotes
    .replace(/\n/g, '\\n')     // Escape newlines
    .replace(/\r/g, '\\r')     // Escape carriage returns
    .replace(/\t/g, '\\t');    // Escape tabs
  
  return text;
}

function escapeHtmlForJson(html) {
  if (!html) return '';
  
  // First escape JSON special characters
  let escaped = html
    .replace(/\\/g, '\\\\')    // Escape backslashes FIRST
    .replace(/"/g, '\\"')      // Escape double quotes
    .replace(/\n/g, '\\n')     // Escape newlines
    .replace(/\r/g, '\\r')     // Escape carriage returns
    .replace(/\t/g, '\\t')     // Escape tabs
    .replace(/\f/g, '\\f');    // Escape form feeds
  
  // Remove control characters
  escaped = escaped.replace(/[\x00-\x1F\x7F-\x9F]/g, '');
  
  return escaped;
}

function testHtmlCleaning() {
  const testHtml = '<span style="font-size: 28px; color: white; text-align: center;"><strong>Docebo\'s AI-based channels may be used to recommend content based on skills and skill levels</strong></span>';
  
  Logger.log('🧪 Testing HTML cleaning...');
  Logger.log(`📋 Original: ${testHtml}`);
  
  const cleaned = stripHtmlTags(testHtml);
  Logger.log(`✅ Cleaned: ${cleaned}`);
  
  const escaped = escapeHtmlForJson(testHtml);
  Logger.log(`🔒 Escaped HTML: ${escaped}`);
  
  // Test JSON creation
  const testJson = `{"content": "${cleaned}"}`;
  Logger.log(`📄 Test JSON: ${testJson}`);
  
  try {
    const parsed = JSON.parse(testJson);
    Logger.log(`✅ JSON parsing successful: ${parsed.content}`);
  } catch (error) {
    Logger.log(`❌ JSON parsing failed: ${error.toString()}`);
  }
}

function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function isValidJsonStructure(jsonString) {
  const trimmed = jsonString.trim();
  
  if (!trimmed.startsWith('{') || !trimmed.endsWith('}')) {
    Logger.log(`❌ JSON doesn't start with { or end with }`);
    return false;
  }
  
  let braceCount = 0;
  let inString = false;
  let escaped = false;
  
  for (let i = 0; i < trimmed.length; i++) {
    const char = trimmed[i];
    
    if (escaped) {
      escaped = false;
      continue;
    }
    
    if (char === '\\') {
      escaped = true;
      continue;
    }
    
    if (char === '"') {
      inString = !inString;
      continue;
    }
    
    if (!inString) {
      if (char === '{') braceCount++;
      if (char === '}') braceCount--;
    }
  }
  
  if (braceCount !== 0) {
    Logger.log(`❌ Unbalanced braces: ${braceCount}`);
    return false;
  }
  
  return true;
}

function extractCoordinateData(htmlContent, parameters) {
  try {
    // Step 1: Extract JSON from HTML template
    const jsonMatch = htmlContent.match(/<script[^>]*id="slide-elements"[^>]*>(.*?)<\/script>/s);
    if (!jsonMatch) {
      Logger.log(`⚠️ No coordinate JSON found in HTML template`);
      return null;
    }
    
    let jsonString = jsonMatch[1].trim();
    Logger.log(`📝 Processing template with ${jsonString.length} characters`);
    
    if (!parameters) {
      Logger.log(`❌ No parameters provided`);
      return null;
    }
    
    Logger.log(`📊 Available data types: ${Object.keys(parameters).join(', ')}`);

    // Step 2: Process all placeholder replacements
    const allReplacements = buildAllReplacements(parameters);
    
    // Step 3: Apply all replacements with HTML cleaning
    Object.entries(allReplacements).forEach(([placeholder, value]) => {
      if (value) {
        const cleanedValue = cleanTextForJson(value);
        const regex = new RegExp(escapeRegExp(placeholder), 'g');
        jsonString = jsonString.replace(regex, cleanedValue);
      }
    });
    
    Logger.log(`✅ All placeholders processed successfully`);

    // Step 4: Validate JSON structure before parsing
    if (!isValidJsonStructure(jsonString)) {
      Logger.log(`❌ JSON structure validation failed`);
      return null;
    }

    // Step 5: Parse JSON and return coordinate data
    const parsedData = JSON.parse(jsonString);
    Logger.log(`✅ Template processed: ${parsedData.elements.length} elements extracted`);
    
    return parsedData;
    
  } catch (error) {
    Logger.log(`❌ Template processing failed: ${error.toString()}`);
    
    // Enhanced error debugging
    const match = error.message.match(/position (\d+)/);
    if (match) {
      const errorPos = parseInt(match[1]);
      Logger.log(`🔍 JSON error at position ${errorPos}`);
      
      // Show error context (simplified)
      const jsonMatch = htmlContent.match(/<script[^>]*id="slide-elements"[^>]*>(.*?)<\/script>/s);
      if (jsonMatch) {
        let debugJsonString = jsonMatch[1].trim();
        const start = Math.max(0, errorPos - 50);
        const end = Math.min(debugJsonString.length, errorPos + 50);
        Logger.log(`🐛 Error context: "${debugJsonString.substring(start, end)}"`);
      }
    }
    
    return null;
  }
}

// Helper function to convert HTML list to array
function convertHtmlListToArray(htmlString) {
  if (!htmlString) return [];
  
  // Extract individual <li> items using regex  
  const liMatches = htmlString.match(/<li[^>]*>.*?<\/li>/gi);
  
  if (liMatches && liMatches.length > 0) {
    const items = liMatches.map(liItem => {
      // Remove all HTML tags from each <li> item
      let cleanText = liItem.replace(/<[^>]*>/g, '');
      
      // Clean up whitespace and normalize
      cleanText = cleanText.replace(/\s+/g, ' ').trim();
      
      // Ensure item ends with period if it doesn't already
      if (cleanText && !cleanText.match(/[.!?]$/)) {
        cleanText += '.';
      }
      
      return cleanText;
    }).filter(item => item.length > 5); // Filter out very short items
    
    return items;
  }
  
  // Fallback: if no <li> tags found, try to split by periods and sentence patterns
  let cleanText = htmlString.replace(/<[^>]*>/g, '');
  let items = cleanText.split(/\.(?=\s*[A-Z])/)
    .filter(item => item.trim().length > 5)
    .map(item => {
      item = item.replace(/\s+/g, ' ').trim();
      if (item && !item.match(/[.!?]$/)) {
        item += '.';
      }
      return item;
    });
  
  return items;
}

function buildAllReplacements(parameters) {
  const replacements = {};
  
  Logger.log(`🔄 Processing ${Object.keys(parameters).length} data types`);
  
  Object.keys(parameters).forEach(dataType => {
    try {
      switch (dataType) {

        case 'doceboTeamArray':
          processTeamMembers(parameters[dataType], replacements, 'docebo');
          break;

        case 'accountTeamArray':
          processTeamMembers(parameters[dataType], replacements, 'account');
          break;
        case 'tableData':
  // This keeps your existing table formatting working if you use type: "table"
  processTableData(parameters[dataType], replacements);
  
  // NEW: Create individual text box mappings for Optimization Strategy fields
  if (Array.isArray(parameters[dataType])) {
    parameters[dataType].forEach((row, index) => {
      const i = index + 1;
      
      // Column 1: Optimization Area
      replacements[`{{area${i}}}`] = row.area || '';
      
      // Column 2: Use Case (Mapped specifically from Optimization Strategy data)
      replacements[`{{optUseCase${i}}}`] = row.useCase || '';
      
      // Column 3: Docebo Feature(s)
      replacements[`{{features${i}}}`] = row.features || '';
      
      // Column 4: Business Impact
      replacements[`{{impact${i}}}`] = row.impact || '';
      
      // Column 5: Mutual Action Plan
      replacements[`{{actionPlan${i}}}`] = row.actionPlan || '';
      
      Logger.log(`✅ Fully mapped Row ${i} from Optimization Strategy`);
    });
  }
  break;
        
        case 'goalsMeet':
          processMeetingGoals(parameters[dataType], replacements);
          break;
        // Add this inside the switch block in buildAllReplacements
        case 'mauContent':
        case 'lsContent':
        case 'ceContent':
          processInsights(parameters, replacements);
          break;
        case 'date':
        case 'logoUrl':
        case 'imageUrl':
        case 'accountName':
        case 'renewalDate':
        case 'strategicValue':
        case 'winsValue':
        case 'productValue':
        case 'lookingValue':
        case 'useCase1':
        case 'useCase2':
        case 'useCase3':
        case 'solutions1':
        case 'solutions2':
        case 'solutions3':
        case 'techStack1':
        case 'techStack2':
        case 'techStack3':
        case 'date1':
        case 'date2':
        case 'date3':
        case 'date4':
        case 'date5':
        case 'key1':
        case 'key2':
        case 'key3':
        case 'key4':
        case 'key5':
        case 'custOwner1':
        case 'custOwner2':
        case 'custOwner3':
        case 'custOwner4':
        case 'custOwner5':
        case 'doceboOwner1':
        case 'doceboOwner2':
        case 'doceboOwner3':
        case 'doceboOwner4':
        case 'doceboOwner5':
        case 'priority1':
        case 'priority2':
        case 'priority3':
        case 'priority4':
        case 'priority5':
        case 'status1':
        case 'status2':
        case 'status3':
        case 'status4':
        case 'status5':
        case 'dueDate1':
        case 'dueDate2':
        case 'dueDate3':
        case 'dueDate4':
        case 'dueDate5':
        case 'milestone1':
        case 'milestone2':
        case 'milestone3':
        case 'milestone4':
        case 'milestone5':
        case 'goal1':
        case 'goal2':
        case 'goal3':
        case 'goal4':
        case 'goal5':
        case 'metrics1':
        case 'metrics2':
        case 'metrics3':
        case 'metrics4':
        case 'metrics5':
        case 'meetGoalOne':
        case 'meetGoalTwo':
        case 'meetGoalThree':
        case 'outcome1':
        case 'outcome2':
        case 'outcome3':
        case 'outcome4':
        case 'outcome5':
        case 'outcome6':
        case 'outcome7':
        case 'outcome8':
        case 'outcome9':
        case 'outcome10':
        case 'outcome11':
        case 'outcome12':
        case 'mauChart':
        case 'lsChart':
        case 'ceChart':
        case 'headerUrl':
        case 'progressImg1':
        case 'slide10Img':
        case 'progressImg3':
        case 'progressImg4':
          processSlideData(parameters, replacements);
          break;
          
        default:
          Logger.log(`⚠️ Unknown data type: ${dataType}`);
          break;
      }
    } catch (error) {
      Logger.log(`❌ Error processing ${dataType}: ${error.toString()}`);
    }
  });
  
  // Process ratio charts after all individual data types are processed
  processRatioCharts(parameters, replacements);
  
  Logger.log(`📝 Built ${Object.keys(replacements).length} placeholder replacements`);
  
  return replacements;
}

function processInsights(parameters, replacements) {
  Logger.log(`📝 Processing insights content`);
  
  // Process mauContent
  const mauContent = parameters.mauContent || '';
  const mauContentArray = convertHtmlListToArray(mauContent);
  replacements['{{mauContent}}'] = JSON.stringify(mauContentArray);
  Logger.log(`✅ Processed mauContent: ${mauContentArray.length} items`);
  
  // Process lsContent
  const lsContent = parameters.lsContent || '';
  const lsContentArray = convertHtmlListToArray(lsContent);
  replacements['{{lsContent}}'] = JSON.stringify(lsContentArray);
  Logger.log(`✅ Processed lsContent: ${lsContentArray.length} items`);
  
  // Process ceContent
  const ceContent = parameters.ceContent || '';
  const ceContentArray = convertHtmlListToArray(ceContent);
  replacements['{{ceContent}}'] = JSON.stringify(ceContentArray);
  Logger.log(`✅ Processed ceContent: ${ceContentArray.length} items`);
}

function processSlideData(parameters, replacements) {
  Logger.log(`📝 Processing simple slide data fields`);

  // Define all simple field names that should be processed
  const simpleFields = [
    'date',
    'logoUrl',
    'imageUrl',
    'accountName',
    'renewalDate',
    'strategicValue',
    'winsValue',
    'productValue',
    'lookingValue',
    'useCase1',
    'useCase2',
    'useCase3',
    'solutions1',
    'solutions2',
    'solutions3',
    'techStack1',
    'techStack2',
    'techStack3',
    'date1',
    'date2',
    'date3',
    'date4',
    'date5',
    'key1',
    'key2',
    'key3',
    'key4',
    'key5',
    'custOwner1',
    'custOwner2',
    'custOwner3',
    'custOwner4',
    'custOwner5',
    'doceboOwner1',
    'doceboOwner2',
    'doceboOwner3',
    'doceboOwner4',
    'doceboOwner5',
    'priority1',
    'priority2',
    'priority3',
    'priority4',
    'priority5',
    'status1',
    'status2',
    'status3',
    'status4',
    'status5',
    'dueDate1',
    'dueDate2',
    'dueDate3',
    'dueDate4',
    'dueDate5',
    'milestone1',
    'milestone2',
    'milestone3',
    'milestone4',
    'milestone5',
    'goal1',
    'goal2',
    'goal3',
    'goal4',
    'goal5',
    'metrics1',
    'metrics2',
    'metrics3',
    'metrics4',
    'metrics5',
    'meetGoalOne',
    'meetGoalTwo',
    'meetGoalThree',
    'outcome1',
    'outcome2',
    'outcome3',
    'outcome4',
    'outcome5',
    'outcome6',
    'outcome7',
    'outcome8',
    'outcome9',
    'outcome10',
    'outcome11',
    'outcome12',
    'mauChart',
    'lsChart',
    'ceChart',
    'headerUrl',
    'progressImg1',
    'slide10Img',
    'progressImg3',
    'progressImg4'
  ];

  let processedCount = 0;

  // Process each field if it exists in parameters
  simpleFields.forEach(fieldName => {
    if (parameters.hasOwnProperty(fieldName)) {
      const value = parameters[fieldName] || '';
      const cleanedValue = cleanTextForJson(value);
      
      // Create replacement with proper placeholder format
      replacements[`{{${fieldName}}}`] = cleanedValue;
      
      processedCount++;
      
      // Log first few for debugging
      if (processedCount <= 5) {
        Logger.log(`  ✅ ${fieldName}: ${cleanedValue.substring(0, 50)}${cleanedValue.length > 50 ? '...' : ''}`);
      }
    }
  });

  Logger.log(`✅ Processed ${processedCount} simple data fields`);
}

// Add this function to process table data
function processTableData(data, replacements) {
  if (!Array.isArray(data)) return;
  
  Logger.log(`📊 Processing tableData (${data.length} rows)`);
  
  // Clean HTML from all table data and convert to 2D array
  const columnKeys = ['area', 'useCase', 'features', 'impact', 'actionPlan'];
  
  // Add header row
  const headerRow = ['Optimization Area', 'Use Case', 'Docebo Feature(s)', 'Business Impact', 'Mutual Action Plan'];
  
  const dataRows = data.map(row => 
    columnKeys.map(key => stripHtmlTags(row[key] || ''))
  );
  
  // Combine header and data rows
  const allRows = [headerRow, ...dataRows];
  
  // Convert to delimited string format
  const rowStrings = allRows.map(row => row.join('|||'));
  const tableString = rowStrings.join('###');
  
  replacements['{{tableData}}'] = tableString;
  
  Logger.log(`✅ Processed ${data.length} table rows + header into delimited string`);
}

function processWhatWeHaveHeard(data, replacements) {
  if (!Array.isArray(data) || data.length === 0) return;
  
  const heardData = data[0];
  Logger.log(`📝 Processing whatWeHaveHeard data`);

  const objectives = heardData.objectives || '';
  const objectivesArray = convertHtmlListToArray(objectives);

  replacements['{{arrayObjectives}}'] = JSON.stringify(objectivesArray);
  
  replacements['{{quote}}'] = heardData.quote || heardData.quote1 || '';
  replacements['{{name}}'] = heardData.name || heardData.name1 || '';
  replacements['{{objectives}}'] = heardData.objectives || '';
  
  replacements['{{quote1}}'] = heardData.quote1 || '';
  replacements['{{quote2}}'] = heardData.quote2 || '';
  replacements['{{quote3}}'] = heardData.quote3 || '';
  replacements['{{name1}}'] = heardData.name1 || '';
  replacements['{{name2}}'] = heardData.name2 || '';
  replacements['{{name3}}'] = heardData.name3 || '';
}

function processMonthlyActiveUsers(data, replacements, fullParameters = null) {
  if (!Array.isArray(data)) return;
  
  Logger.log(`📊 Processing monthlyActiveUsers data (${data.length} records)`);

  // Sort data by date for calculations (newest first for getting latest values)
  const sortedDataNewestFirst = [...data].sort((a, b) => new Date(b.month + " 1") - new Date(a.month + " 1"));

  // Sort data by date for chart display (oldest first for chronological chart)
  const sortedDataOldestFirst = [...data].sort((a, b) => new Date(a.month + " 1") - new Date(b.month + " 1"));

  // Get plan users data for overage calculations
  let planUsersData = [];
  if (fullParameters && fullParameters.monthlyPlanUsers && Array.isArray(fullParameters.monthlyPlanUsers)) {
    planUsersData = fullParameters.monthlyPlanUsers;
    Logger.log(`📊 Found monthlyPlanUsers data (${planUsersData.length} records)`);
  } else {
    Logger.log(`⚠️ No monthlyPlanUsers data found for overage calculations`);
  }

  // Create chart data format - using oldest first for proper chart display
  const chartData = {
    "columns": [
      {"type": "string", "label": "Month"},
      {"type": "number", "label": "Users"}
    ],
    "rows": sortedDataOldestFirst.map(item => [item.month || '', item.count || 0])
  };

  Logger.log(`Chart data created with ${chartData.rows.length} rows (oldest to newest)`);
  Logger.log(`Chart first row (oldest): ${JSON.stringify(chartData.rows[0])}`);
  Logger.log(`Chart last row (newest): ${JSON.stringify(chartData.rows[chartData.rows.length - 1])}`);

  // Calculate counter metrics using newest first data
  const latestCount = sortedDataNewestFirst[0]?.count || 0;
  const previousCount = sortedDataNewestFirst[1]?.count || 0;
  
  // Calculate YoY change (comparing same month from previous year)
  let yoyChange = 0;
  let yoyPercentage = "0%";
  
  if (sortedDataNewestFirst.length > 0) {
    const latestMonth = sortedDataNewestFirst[0].month;
    const latestYear = new Date(latestMonth + " 1").getFullYear();
    const targetMonth = latestMonth.replace(latestYear.toString(), (latestYear - 1).toString());
    
    const previousYearData = sortedDataNewestFirst.find(item => item.month === targetMonth);
    if (previousYearData) {
      yoyChange = latestCount - previousYearData.count;
      yoyPercentage = previousYearData.count > 0 
        ? Math.round((yoyChange / previousYearData.count) * 100) + "%" 
        : "0%";
    }
  }
  
  // Calculate overage metrics using plan users data
  let maxOverage = 0;
  let periodsWithOverages = 0;
  let totalOverageUsers = 0;
  let overageValues = [];
  
  if (planUsersData.length > 0) {
    // Calculate overages for each month
    sortedDataNewestFirst.forEach(activeItem => {
      const planItem = planUsersData.find(p => p.month === activeItem.month);
      if (planItem) {
        const activeUsers = activeItem.count || 0;
        const planUsers = planItem.count || 0;
        const overage = activeUsers - planUsers;
        
        if (overage > 0) {
          periodsWithOverages++;
          totalOverageUsers += overage;
          overageValues.push(overage);
          maxOverage = Math.max(maxOverage, overage);
        }
        
        Logger.log(`  ${activeItem.month}: Active=${activeUsers}, Plan=${planUsers}, Overage=${overage}`);
      }
    });
  } else {
    Logger.log(`⚠️ Cannot calculate overage metrics without plan users data`);
  }
  
  // Calculate average overage users (only for periods with actual overages)
  const avgOverageUsers = overageValues.length > 0 
    ? Math.round(totalOverageUsers / overageValues.length) 
    : 0;

  Logger.log(`Calculated metrics:`);
  Logger.log(`  Last month active users: ${latestCount}`);
  Logger.log(`  YoY change: ${yoyChange} (${yoyPercentage})`);
  Logger.log(`  Max overage: ${maxOverage}`);
  Logger.log(`  Periods with overages: ${periodsWithOverages}`);
  Logger.log(`  Avg overage users: ${avgOverageUsers}`);

  // Store different formats for flexibility
  replacements['{{monthlyActiveUsersData}}'] = JSON.stringify(data); // Original format
  replacements['{{monthlyActiveUsersChartData}}'] = JSON.stringify(chartData); // Chart format (oldest to newest)
  
  // Counter metrics (using newest first data)
  replacements['{{latestActiveUsers}}'] = latestCount.toString();
  replacements['{{latestActiveUsersMonth}}'] = sortedDataNewestFirst[0]?.month || '';
  replacements['{{yoyMAUsChange}}'] = yoyChange.toString();
  replacements['{{yoyMAUsPercentage}}'] = yoyPercentage;
  replacements['{{maxOverage}}'] = maxOverage.toString(); // Fixed calculation
  replacements['{{periodsWithOverages}}'] = periodsWithOverages.toString();
  replacements['{{avgOverageUsers}}'] = avgOverageUsers.toString(); // Fixed calculation
  
  // Additional useful metrics
  replacements['{{previousMonthActiveUsers}}'] = previousCount.toString();
  replacements['{{monthlyChange}}'] = (latestCount - previousCount).toString();
  replacements['{{monthlyChangePercentage}}'] = previousCount > 0 
    ? Math.round(((latestCount - previousCount) / previousCount) * 100) + "%" 
    : "0%";
  
  // Also store raw chart data (without JSON.stringify) for direct use
  replacements['{{monthlyActiveUsersChartDataRaw}}'] = chartData;
}


function processMonthlyRegisteredUsers(data, replacements) {
  if (!Array.isArray(data)) return;
  
  Logger.log(`📊 Processing monthlyRegisteredUsers data (${data.length} records)`);

  // Sort data by date for calculations (newest first for getting latest values)
  const sortedDataNewestFirst = [...data].sort((a, b) => new Date(b.month + " 1") - new Date(a.month + " 1"));

  // Sort data by date for chart display (oldest first for chronological chart)
  const sortedDataOldestFirst = [...data].sort((a, b) => new Date(a.month + " 1") - new Date(b.month + " 1"));
  
  // Create chart data format - using oldest first for proper chart display
  const chartData = {
    "columns": [
      {"type": "string", "label": "Month"},
      {"type": "number", "label": "Users"}
    ],
    "rows": sortedDataOldestFirst.map(item => [item.month || '', item.value || 0])
  };

  Logger.log(`Chart data created with ${chartData.rows.length} rows (oldest to newest)`);
  Logger.log(`Chart first row (oldest): ${JSON.stringify(chartData.rows[0])}`);
  Logger.log(`Chart last row (newest): ${JSON.stringify(chartData.rows[chartData.rows.length - 1])}`);

  // Calculate counter metrics using newest first data
  const latestCount = sortedDataNewestFirst[0]?.value || 0;
  const previousCount = sortedDataNewestFirst[1]?.value || 0;
  
  // Calculate YoY change (comparing same month from previous year)
  let yoyChange = 0;
  let yoyPercentage = "0%";
  
  if (sortedDataNewestFirst.length > 0) {
    const latestMonth = sortedDataNewestFirst[0].month;
    const latestYear = new Date(latestMonth + " 1").getFullYear();
    const targetMonth = latestMonth.replace(latestYear.toString(), (latestYear - 1).toString());
    
    const previousYearData = sortedDataNewestFirst.find(item => item.month === targetMonth);
    if (previousYearData) {
      yoyChange = latestCount - previousYearData.value;
      yoyPercentage = previousYearData.value > 0 
        ? Math.round((yoyChange / previousYearData.value) * 100) + "%" 
        : "0%";
    }
  }

  Logger.log(`Calculated registered users metrics:`);
  Logger.log(`  Last month registered users: ${latestCount}`);
  Logger.log(`  YoY change: ${yoyChange} (${yoyPercentage})`);
  Logger.log(`  Previous month: ${previousCount}`);

  // Store different formats for flexibility - using newest first data for latest values
  replacements['{{monthlyRegisteredUsersData}}'] = JSON.stringify(data); // Original format
  replacements['{{monthlyRegisteredUsersChartData}}'] = JSON.stringify(chartData); // Chart format (oldest to newest)
  
  // Counter metrics
  replacements['{{latestRegisteredUsers}}'] = latestCount.toString(); // Last month registered users
  replacements['{{latestRegisteredUsersMonth}}'] = sortedDataNewestFirst[0]?.month || '';
  replacements['{{yoyRUsChange}}'] = yoyChange.toString(); // YoY change in absolute numbers
  replacements['{{yoyRUsPercentage}}'] = yoyPercentage; // YoY ΔRUs percentage
  
  // Additional useful metrics
  replacements['{{previousMonthRegisteredUsers}}'] = previousCount.toString();
  replacements['{{monthlyRegisteredChange}}'] = (latestCount - previousCount).toString();
  replacements['{{monthlyRegisteredChangePercentage}}'] = previousCount > 0 
    ? Math.round(((latestCount - previousCount) / previousCount) * 100) + "%" 
    : "0%";
  
  // Also store raw chart data (without JSON.stringify) for direct use
  replacements['{{monthlyRegisteredUsersChartDataRaw}}'] = chartData;
}

function processMonthlyPlanUsers(data, replacements) {
  if (!Array.isArray(data)) return;
  
  Logger.log(`📊 Processing monthlyPlanUsers data (${data.length} records)`);

  // Sort data by date for calculations (newest first for getting latest values)
  const sortedDataNewestFirst = [...data].sort((a, b) => new Date(b.month + " 1") - new Date(a.month + " 1"));

  // Sort data by date for chart display (oldest first for chronological chart)
  const sortedDataOldestFirst = [...data].sort((a, b) => new Date(a.month + " 1") - new Date(b.month + " 1"));

  // Create chart data format - using oldest first for proper chart display
  const chartData = {
    "columns": [
      {"type": "string", "label": "Month"},
      {"type": "number", "label": "Plan Users"}
    ],
    "rows": sortedDataOldestFirst.map(item => [item.month || '', item.count || 0])
  };

  Logger.log(`Plan users chart data created with ${chartData.rows.length} rows (oldest to newest)`);
  Logger.log(`Chart first row (oldest): ${JSON.stringify(chartData.rows[0])}`);
  Logger.log(`Chart last row (newest): ${JSON.stringify(chartData.rows[chartData.rows.length - 1])}`);

  // Store different formats for flexibility - using newest first data for latest values
  replacements['{{monthlyPlanUsersData}}'] = JSON.stringify(data); // Original format
  replacements['{{monthlyPlanUsersChartData}}'] = JSON.stringify(chartData); // Chart format (oldest to newest)
  replacements['{{latestPlanUsers}}'] = sortedDataNewestFirst[0]?.count || ''; // Using newest first for latest value
  replacements['{{latestPlanUsersMonth}}'] = sortedDataNewestFirst[0]?.month || '';
  
  // Also store raw chart data (without JSON.stringify) for direct use
  replacements['{{monthlyPlanUsersChartDataRaw}}'] = chartData;
}

// Process Active vs Registered Users Ratio chart
function processActiveVsRegisteredRatio(parameters, replacements) {
  const activeUsersData = parameters.monthlyActiveUsers || [];
  const registeredUsersData = parameters.monthlyRegisteredUsers || [];
  
  if (!Array.isArray(activeUsersData) || !Array.isArray(registeredUsersData)) return;
  
  Logger.log(`📊 Processing Active vs Registered Users Ratio`);
  
  // Sort both datasets by date (oldest first for chart display)
  const sortedActiveData = [...activeUsersData].sort((a, b) => new Date(a.month + " 1") - new Date(b.month + " 1"));
  const sortedRegisteredData = [...registeredUsersData].sort((a, b) => new Date(a.month + " 1") - new Date(b.month + " 1"));
  
  // Combine data for dual-axis chart (bars + line)
  const combinedData = [];
  sortedActiveData.forEach(activeItem => {
    const registeredItem = sortedRegisteredData.find(r => r.month === activeItem.month);
    if (registeredItem) {
      combinedData.push([
        activeItem.month || '',
        activeItem.count || 0,        // Active Users (for line)
        registeredItem.value || 0     // Registered Users (for bars) - note: using 'value' not 'count'
      ]);
    }
  });
  
  const activeVsRegisteredChartData = {
    "columns": [
      {"type": "string", "label": "Month"},
      {"type": "number", "label": "Active Users"},
      {"type": "number", "label": "Registered Users"}
    ],
    "rows": combinedData
  };
  
  Logger.log(`Active vs Registered chart data: ${combinedData.length} rows`);
  
  replacements['{{activeVsRegisteredRatioChartData}}'] = JSON.stringify(activeVsRegisteredChartData);
  replacements['{{activeVsRegisteredRatioChartDataRaw}}'] = activeVsRegisteredChartData;
}

// Process Learning Sessions per Active User chart
function processLearningSessionsPerUser(parameters, replacements) {
  const timeSpentData = parameters.monthlyTimeSpentInLearning || [];
  const activeUsersData = parameters.monthlyActiveUsers || [];
  
  if (!Array.isArray(timeSpentData) || !Array.isArray(activeUsersData)) return;
  
  Logger.log(`📊 Processing Learning Sessions per Active User`);
  
  // Helper function to convert minutes to hours
  const minutesToHours = (minutes) => {
    return Math.round((minutes / 60) * 100) / 100;
  };
  
  // Sort both datasets by date (oldest first for chart display)
  const sortedTimeData = [...timeSpentData].sort((a, b) => new Date(a.month + " 1") - new Date(b.month + " 1"));
  const sortedActiveData = [...activeUsersData].sort((a, b) => new Date(a.month + " 1") - new Date(b.month + " 1"));
  
  // Combine data for dual-axis chart
  const combinedData = [];
  sortedTimeData.forEach(timeItem => {
    const activeItem = sortedActiveData.find(a => a.month === timeItem.month);
    if (activeItem) {
      const learningHours = minutesToHours(timeItem["Time In Min"] || 0);
      const activeUsers = activeItem.count || 0;
      
      combinedData.push([
        timeItem.month || '',
        learningHours,      // Learning Sessions (hours) for bars
        activeUsers         // Active Users for line
      ]);
    }
  });
  
  const learningSessionsPerUserChartData = {
    "columns": [
      {"type": "string", "label": "Month"},
      {"type": "number", "label": "Learning Hours"},
      {"type": "number", "label": "Active Users"}
    ],
    "rows": combinedData
  };
  
  Logger.log(`Learning Sessions per User chart data: ${combinedData.length} rows`);
  
  replacements['{{learningSessionsPerUserChartData}}'] = JSON.stringify(learningSessionsPerUserChartData);
  replacements['{{learningSessionsPerUserChartDataRaw}}'] = learningSessionsPerUserChartData;
}

// Process Monthly Usage Rate (MAUs : Plan Ratio) - Horizontal Bar Chart
function processMonthlyUsageRate(parameters, replacements) {
  const activeUsersData = parameters.monthlyActiveUsers || [];
  const planUsersData = parameters.monthlyPlanUsers || [];
  
  if (!Array.isArray(activeUsersData) || !Array.isArray(planUsersData)) return;
  
  Logger.log(`📊 Processing Monthly Usage Rate (MAUs : Plan Ratio)`);
  
  // Sort both datasets by date (newest first for this chart to show recent months at top)
  const sortedActiveData = [...activeUsersData].sort((a, b) => new Date(b.month + " 1") - new Date(a.month + " 1"));
  const sortedPlanData = [...planUsersData].sort((a, b) => new Date(b.month + " 1") - new Date(a.month + " 1"));
  
  // Calculate usage ratios
  const usageRateData = [];
  sortedActiveData.forEach(activeItem => {
    const planItem = sortedPlanData.find(p => p.month === activeItem.month);
    if (planItem) {
      const activeUsers = activeItem.count || 0;
      const planUsers = planItem.count || 0;
      const usagePercentage = planUsers > 0 ? Math.round((activeUsers / planUsers) * 100) : 0;
      
      // For horizontal bar chart, we need both active and plan values
      usageRateData.push([
        activeItem.month || '',
        activeUsers,      // Active Users
        planUsers        // Plan Users (full capacity)
      ]);
    }
  });
  
  const monthlyUsageRateChartData = {
    "columns": [
      {"type": "string", "label": "Month"},
      {"type": "number", "label": "Active Users"},
      {"type": "number", "label": "Plan Users"}
    ],
    "rows": usageRateData
  };
  
  Logger.log(`Monthly Usage Rate chart data: ${usageRateData.length} rows`);
  Logger.log(`Sample usage data: ${JSON.stringify(usageRateData[0])}`);
  
  replacements['{{monthlyUsageRateChartData}}'] = JSON.stringify(monthlyUsageRateChartData);
  replacements['{{monthlyUsageRateChartDataRaw}}'] = monthlyUsageRateChartData;
}

// Main function to process all ratio charts
function processRatioCharts(parameters, replacements) {
  Logger.log(`📊 Processing all ratio charts`);
  
  processActiveVsRegisteredRatio(parameters, replacements);
  processLearningSessionsPerUser(parameters, replacements);
  processMonthlyUsageRate(parameters, replacements);
}

function processMonthlyTimeSpentInLearning(data, replacements) {
  if (!Array.isArray(data)) return;
  
  Logger.log(`📊 Processing monthlyTimeSpentInLearning data (${data.length} records)`);
  
  // Helper function to convert minutes to hours
  const minutesToHours = (minutes) => {
    return Math.round((minutes / 60) * 100) / 100; // Round to 2 decimal places
  };
  
  // Sort data by date for calculations (newest first for getting latest values)
  const sortedDataNewestFirst = [...data].sort((a, b) => new Date(b.month + " 1") - new Date(a.month + " 1"));

  // Sort data by date for chart display (oldest first for chronological chart)
  const sortedDataOldestFirst = [...data].sort((a, b) => new Date(a.month + " 1") - new Date(b.month + " 1"));
  
  // Check if data has time values or just months
  const hasTimeData = sortedDataNewestFirst.some(item => 
    item["Time In Min"] || item.timeInMin || 
    item["Ilt Time In Min"] || item.iltTimeInMin || 
    item["E Learning Time In Min"] || item.eLearningTimeInMin
  );
  
  if (!hasTimeData) {
    Logger.log(`⚠️ No time data found in monthlyTimeSpentInLearning. Using zero values.`);
    Logger.log(`Expected properties: "Time In Min", "Ilt Time In Min", "E Learning Time In Min"`);
  }
  
  // Transform data for Learning Sessions chart (column chart) - using oldest first for proper chart display
  const learningSessionsData = sortedDataOldestFirst.map(item => {
    const totalTimeMinutes = (item["Time In Min"] || item.timeInMin || 0);
    const totalTimeHours = minutesToHours(totalTimeMinutes);
    
    return [item.month || '', totalTimeHours];
  });
  
  const learningSessionsChartData = {
    "columns": [
      {"type": "string", "label": "Month"},
      {"type": "number", "label": "Hours"}
    ],
    "rows": learningSessionsData
  };
  
  // Transform data for Sessions by Type chart (line chart with ILT and eLearning) - using oldest first for proper chart display
  const sessionsByTypeData = sortedDataOldestFirst.map(item => {
    const iltTimeMinutes = (item["Ilt Time In Min"] || item.iltTimeInMin || 0);
    const eLearningTimeMinutes = (item["E Learning Time In Min"] || item.eLearningTimeInMin || 0);
    
    const iltTimeHours = minutesToHours(iltTimeMinutes);
    const eLearningTimeHours = minutesToHours(eLearningTimeMinutes);
    
    return [item.month || '', iltTimeHours, eLearningTimeHours];
  });
  
  const sessionsByTypeChartData = {
    "columns": [
      {"type": "string", "label": "Month"},
      {"type": "number", "label": "ILT"},
      {"type": "number", "label": "eLearning"}
    ],
    "rows": sessionsByTypeData
  };
  
  Logger.log(`Learning Sessions chart data: ${learningSessionsData.length} rows (oldest to newest)`);
  Logger.log(`Sessions by Type chart data: ${sessionsByTypeData.length} rows (oldest to newest)`);
  Logger.log(`Chart first row (oldest): ${JSON.stringify(learningSessionsData[0])}`);
  Logger.log(`Chart last row (newest): ${JSON.stringify(learningSessionsData[learningSessionsData.length - 1])}`);
  
  // Store original data
  replacements['{{monthlyTimeSpentData}}'] = JSON.stringify(data);
  
  // Store chart-ready data for Learning Sessions (column chart)
  replacements['{{learningSessionsChartData}}'] = JSON.stringify(learningSessionsChartData);
  replacements['{{learningSessionsChartDataRaw}}'] = learningSessionsChartData;
  
  // Store chart-ready data for Sessions by Type (line chart)
  replacements['{{sessionsByTypeChartData}}'] = JSON.stringify(sessionsByTypeChartData);
  replacements['{{sessionsByTypeChartDataRaw}}'] = sessionsByTypeChartData;
  
  // Store latest values (using newest first data)
  if (sortedDataNewestFirst.length > 0) {
    const latestData = sortedDataNewestFirst[0];
    replacements['{{latestTotalTimeHours}}'] = minutesToHours(latestData["Time In Min"] || latestData.timeInMin || 0).toString();
    replacements['{{latestIltTimeHours}}'] = minutesToHours(latestData["Ilt Time In Min"] || latestData.iltTimeInMin || 0).toString();
    replacements['{{latestELearningTimeHours}}'] = minutesToHours(latestData["E Learning Time In Min"] || latestData.eLearningTimeInMin || 0).toString();
    replacements['{{latestMonth}}'] = latestData.month || '';
  }
}


function processCourseEnrollment(data, replacements) {
  if (!Array.isArray(data) || data.length === 0) {
    Logger.log(`⚠️ No courseEnrollment data available`);
    replacements['{{courseEnrollmentData}}'] = JSON.stringify([]);
    replacements['{{courseTypePieChartData}}'] = JSON.stringify({columns: [], rows: []});
    replacements['{{courseProviderPieChartData}}'] = JSON.stringify({columns: [], rows: []});
    replacements['{{courseDistributionChartData}}'] = JSON.stringify({columns: [], rows: []});
    return;
  }
  
  Logger.log(`📊 Processing courseEnrollment data (${data.length} records)`);
  
  // Check if we have meaningful enrollment data
  const hasMeaningfulData = data.some(item => 
    (item["E Learning Count"] && item["E Learning Count"] > 0) || 
    (item["Ilt Count"] && item["Ilt Count"] > 0)
  );
  
  if (!hasMeaningfulData) {
    Logger.log(`⚠️ No meaningful enrollment counts found in data. Skipping chart generation.`);
    replacements['{{courseEnrollmentData}}'] = JSON.stringify(data);
    replacements['{{courseTypePieChartData}}'] = JSON.stringify({columns: [], rows: []});
    replacements['{{courseProviderPieChartData}}'] = JSON.stringify({columns: [], rows: []});
    replacements['{{courseDistributionChartData}}'] = JSON.stringify({columns: [], rows: []});
    return;
  }
  
  // Sort data by date (oldest first for chart display)
  const sortedDataOldestFirst = [...data].sort((a, b) => new Date(a.month + " 1") - new Date(b.month + " 1"));
  
  // Calculate totals for pie charts
  let totalELearning = 0;
  let totalILT = 0;
  let totalInternal = 0;
  let totalDocebo = 0;
  let totalAdmin = 0;
  let totalSelf = 0;
  
  // Sum up all enrollment counts across all months
  data.forEach(item => {
    totalELearning += item["E Learning Count"] || 0;
    totalILT += item["Ilt Count"] || 0;
    totalInternal += item["Internal Count"] || 0;
    totalDocebo += item["Docebo Count"] || 0;
    totalAdmin += item["Admin Count"] || 0;
    totalSelf += item["Self Count"] || 0;
  });
  
  // Create Course Type Pie Chart data (E Learning vs ILT)
  const courseTypePieData = {
    "columns": [
      {"type": "string", "label": "Course Type"},
      {"type": "number", "label": "Enrollments"}
    ],
    "rows": [
      ["E Learning", totalELearning],
      ["ILT", totalILT]
    ]
  };
  
  // Create Course Provider Pie Chart data (Internal vs Docebo vs Others)
  const courseProviderData = [];
  if (totalInternal > 0) courseProviderData.push(["Internal", totalInternal]);
  if (totalDocebo > 0) courseProviderData.push(["Docebo", totalDocebo]); 
  if (totalAdmin > 0) courseProviderData.push(["Admin", totalAdmin]);
  if (totalSelf > 0) courseProviderData.push(["Self", totalSelf]);
  
  // If no provider data, show 100% Internal as fallback
  if (courseProviderData.length === 0) {
    courseProviderData.push(["Internal", 100]);
  }
  
  const courseProviderPieData = {
    "columns": [
      {"type": "string", "label": "Provider"},
      {"type": "number", "label": "Enrollments"}
    ],
    "rows": courseProviderData
  };
  
  // Create Course Type Distribution Column Chart data (E Learning vs ILT over time)
  const courseDistributionData = sortedDataOldestFirst.map(item => [
    item.month || '',
    item["E Learning Count"] || 0,
    item["Ilt Count"] || 0
  ]);
  
  const courseDistributionChartData = {
    "columns": [
      {"type": "string", "label": "Month"},
      {"type": "number", "label": "E Learning"},
      {"type": "number", "label": "ILT"}
    ],
    "rows": courseDistributionData
  };
  
  Logger.log(`Course enrollment totals: E Learning=${totalELearning}, ILT=${totalILT}`);
  Logger.log(`Course provider totals: Internal=${totalInternal}, Docebo=${totalDocebo}`);
  Logger.log(`Course distribution chart: ${courseDistributionData.length} months`);
  Logger.log(`Sample distribution row: ${JSON.stringify(courseDistributionData[0])}`);
  
  // Store original data
  replacements['{{courseEnrollmentData}}'] = JSON.stringify(data);
  
  // Store chart-ready data
  replacements['{{courseTypePieChartData}}'] = JSON.stringify(courseTypePieData);
  replacements['{{courseProviderPieChartData}}'] = JSON.stringify(courseProviderPieData);
  replacements['{{courseDistributionChartData}}'] = JSON.stringify(courseDistributionChartData);
  
  // Store raw chart data (without JSON.stringify) for direct use
  replacements['{{courseTypePieChartDataRaw}}'] = courseTypePieData;
  replacements['{{courseProviderPieChartDataRaw}}'] = courseProviderPieData;
  replacements['{{courseDistributionChartDataRaw}}'] = courseDistributionChartData;
  
  // Store summary metrics
  replacements['{{totalELearningEnrollments}}'] = totalELearning.toString();
  replacements['{{totalILTEnrollments}}'] = totalILT.toString();
  replacements['{{totalEnrollments}}'] = (totalELearning + totalILT).toString();
  
  // Calculate percentages
  const totalCourseEnrollments = totalELearning + totalILT;
  const eLearningPercentage = totalCourseEnrollments > 0 ? Math.round((totalELearning / totalCourseEnrollments) * 100) : 0;
  const iltPercentage = totalCourseEnrollments > 0 ? Math.round((totalILT / totalCourseEnrollments) * 100) : 0;
  
  replacements['{{eLearningPercentage}}'] = eLearningPercentage.toString() + '%';
  replacements['{{iltPercentage}}'] = iltPercentage.toString() + '%';
  
  // Latest month data
  const latestData = [...data].sort((a, b) => new Date(b.month + " 1") - new Date(a.month + " 1"))[0];
  if (latestData) {
    replacements['{{latestELearningEnrollments}}'] = (latestData["E Learning Count"] || 0).toString();
    replacements['{{latestILTEnrollments}}'] = (latestData["Ilt Count"] || 0).toString();
    replacements['{{latestEnrollmentMonth}}'] = latestData.month || '';
  }
}


function processCourseCompletion(data, replacements, fullParameters = null) {
  if (!Array.isArray(data) || data.length === 0) {
    Logger.log(`⚠️ No courseCompletion data available`);
    replacements['{{courseCompletionData}}'] = JSON.stringify([]);
    replacements['{{courseCompletionChartData}}'] = JSON.stringify({columns: [], rows: []});
    replacements['{{totalCourseCompletions}}'] = "0";
    replacements['{{completionRate}}'] = "0%";
    return;
  }
  
  Logger.log(`📊 Processing courseCompletion data (${data.length} records)`);
  
  // Sort data by date (oldest first for chart display)
  const sortedDataOldestFirst = [...data].sort((a, b) => new Date(a.month + " 1") - new Date(b.month + " 1"));
  
  // Check if we have actual completion data with counts
  const hasCompletionCounts = data.some(item => 
    (item["E Learning Count"] && item["E Learning Count"] > 0) || 
    (item["Ilt Count"] && item["Ilt Count"] > 0) || 
    (item["Total Count"] && item["Total Count"] > 0)
  );
  
  if (!hasCompletionCounts) {
    Logger.log(`⚠️ No meaningful completion counts found in data. Skipping chart generation.`);
    replacements['{{courseCompletionData}}'] = JSON.stringify(data);
    replacements['{{courseCompletionChartData}}'] = JSON.stringify({columns: [], rows: []});
    replacements['{{totalCourseCompletions}}'] = "0";
    replacements['{{completionRate}}'] = "0%";
    return;
  }
  
  Logger.log(`✅ Found actual completion data with counts`);
  
  // Calculate completion metrics using actual data
  let totalCompletions = 0;
  let totalELearningCompletions = 0;
  let totalILTCompletions = 0;
  
  // Sum up all completion counts across all months
  data.forEach(item => {
    const eLearningComps = item["E Learning Count"] || 0;
    const iltComps = item["Ilt Count"] || 0;
    const totalComps = item["Total Count"] || (eLearningComps + iltComps);
    
    totalELearningCompletions += eLearningComps;
    totalILTCompletions += iltComps;
    totalCompletions += totalComps;
  });
  
  // Create chart data for course completions over time
  const courseCompletionChartData = sortedDataOldestFirst.map(item => {
    const eLearningComps = item["E Learning Count"] || 0;
    const iltComps = item["Ilt Count"] || 0;
    
    return [item.month || '', eLearningComps, iltComps];
  });
  
  const courseCompletionTimelineData = {
    "columns": [
      {"type": "string", "label": "Month"},
      {"type": "number", "label": "eLearning"},
      {"type": "number", "label": "ILT"}
    ],
    "rows": courseCompletionChartData
  };
  
  // Calculate completion rate using enrollment data if available
  let completionRate = 0;
  let totalEnrollments = 0;
  
  if (fullParameters && fullParameters.courseEnrollment && Array.isArray(fullParameters.courseEnrollment)) {
    const enrollmentData = fullParameters.courseEnrollment;
    
    // Calculate total enrollments from enrollment data
    totalEnrollments = enrollmentData.reduce((sum, item) => 
      sum + (item["E Learning Count"] || 0) + (item["Ilt Count"] || 0), 0
    );
    
    if (totalEnrollments > 0) {
      completionRate = Math.round((totalCompletions / totalEnrollments) * 100);
      Logger.log(`📊 Calculated completion rate: ${totalCompletions} completions / ${totalEnrollments} enrollments = ${completionRate}%`);
    }
  } else {
    Logger.log(`⚠️ No enrollment data available for completion rate calculation`);
    // If no enrollment data, we can't calculate a meaningful completion rate
    completionRate = 0;
  }
  
  Logger.log(`Course completion metrics:`);
  Logger.log(`  Total completions: ${totalCompletions}`);
  Logger.log(`  E Learning completions: ${totalELearningCompletions}`);
  Logger.log(`  ILT completions: ${totalILTCompletions}`);
  Logger.log(`  Completion rate: ${completionRate}%`);
  Logger.log(`  Chart data: ${courseCompletionChartData.length} months`);
  
  // Store original data
  replacements['{{courseCompletionData}}'] = JSON.stringify(data);
  
  // Store chart data
  replacements['{{courseCompletionChartData}}'] = JSON.stringify(courseCompletionTimelineData);
  replacements['{{courseCompletionChartDataRaw}}'] = courseCompletionTimelineData;
  
  // Store completion metrics
  replacements['{{totalCourseCompletions}}'] = totalCompletions.toString();
  replacements['{{totalELearningCompletions}}'] = totalELearningCompletions.toString();
  replacements['{{totalILTCompletions}}'] = totalILTCompletions.toString();
  replacements['{{completionRate}}'] = completionRate.toString() + '%';
  replacements['{{completionRateNumber}}'] = completionRate.toString();
  
  // Latest month data
  const latestData = [...data].sort((a, b) => new Date(b.month + " 1") - new Date(a.month + " 1"))[0];
  if (latestData) {
    replacements['{{latestCompletionMonth}}'] = latestData.month || '';
    replacements['{{latestELearningCompletions}}'] = (latestData["E Learning Count"] || 0).toString();
    replacements['{{latestILTCompletions}}'] = (latestData["Ilt Count"] || 0).toString();
    replacements['{{latestTotalCompletions}}'] = (latestData["Total Count"] || 0).toString();
  }
}

function processTeamMembers(data, replacements, teamType) {
  if (!Array.isArray(data)) {
    Logger.log(`⚠️ Team members data is not an array`);
    return;
  }
  
  Logger.log(`👥 Processing ${teamType || 'team'} members data (${data.length} members)`);
  
  // Process the team members array
  const processedMembers = data.map(member => ({
    profilePicture: member.profilePicture || "Unavailable",
    name: stripHtmlTags(member.name || ''),
    membership: member.membership || ''
  }));
  
  Logger.log(`✅ Processed ${processedMembers.length} ${teamType || 'team'} members`);
  
  // Determine placeholder names based on team type
  if (teamType === 'docebo') {
    replacements['{{doceboTeamArray}}'] = JSON.stringify(processedMembers);
    replacements['{{doceboTeamArrayRaw}}'] = processedMembers;
    replacements['{{doceboTeamCount}}'] = processedMembers.length.toString();
  } else if (teamType === 'account') {
    replacements['{{accountTeamArray}}'] = JSON.stringify(processedMembers);
    replacements['{{accountTeamArrayRaw}}'] = processedMembers;
    replacements['{{accountTeamCount}}'] = processedMembers.length.toString();
  }
}

function processWins(data, replacements) {
  if (!Array.isArray(data)) return;
  
  Logger.log(`🏆 Processing wins data (${data.length} items)`);
  
  replacements['{{wins}}'] = JSON.stringify(data);
  replacements['{{winsCount}}'] = data.length.toString();
  
  data.forEach((win, index) => {
    replacements[`{{win${index + 1}Title}}`] = win.title || '';
    replacements[`{{win${index + 1}Summary}}`] = win.summary || '';
    replacements[`{{win${index + 1}KPI}}`] = win.relatedKPI || '';
  });
}

function processOpportunities(data, replacements) {
  if (!Array.isArray(data)) return;
  
  Logger.log(`🎯 Processing opportunities data (${data.length} items)`);
  
  replacements['{{opportunities}}'] = JSON.stringify(data);
  replacements['{{opportunitiesCount}}'] = data.length.toString();
  
  data.forEach((opp, index) => {
    replacements[`{{opportunity${index + 1}Description}}`] = opp.description || '';
    replacements[`{{opportunity${index + 1}Result}}`] = opp.result || '';
  });
}

function processStrategicObjectives(data, replacements) {
  if (!Array.isArray(data)) return;
  
  Logger.log(`📋 Processing strategic objectives data (${data.length} items)`);
  
  replacements['{{strategicObjectives}}'] = JSON.stringify(data);
  
  // Create individual placeholders for each strategic objective by name
  data.forEach(item => {
    const key = item.name?.replace(/\s+/g, '').replace(/[^a-zA-Z0-9]/g, '');
    if (key) {
      replacements[`{{${key}}}`] = item.value || '';
      Logger.log(`   Created placeholder: {{${key}}}`);
    }
  });
}

function convertToChartData(data, labelField, valueField) {
  if (!Array.isArray(data) || data.length === 0) {
    return { columns: [], rows: [] };
  }
  
  const columns = [
    { label: labelField.charAt(0).toUpperCase() + labelField.slice(1) },
    { label: valueField.charAt(0).toUpperCase() + valueField.slice(1) }
  ];
  
  const rows = data.map(item => [
    item[labelField] || '',
    item[valueField] || 0
  ]);
  
  return { columns, rows };
}

function calculateGrowth(data, valueField) {
  if (!Array.isArray(data) || data.length < 2) {
    return '0%';
  }
  
  const current = data[0][valueField] || 0;
  const previous = data[1][valueField] || 0;
  
  if (previous === 0) return '0%';
  
  const growth = ((current - previous) / previous * 100).toFixed(1);
  return `${growth > 0 ? '+' : ''}${growth}%`;
}

function renderCoordinateElements(slide, data, slideName, presentation) {
  try {
    Logger.log(`🎨 Rendering ${data.elements.length} elements with exact coordinates...`);
    
    data.elements.forEach((element, index) => {
      try {
        // Pass slideName for chart naming
        renderCoordinateElement(slide, element, slideName, presentation);
      } catch (elementError) {
        Logger.log(`⚠️ Error rendering element ${index}: ${elementError.toString()}`);
      }
    });
    
    Logger.log(`✅ All coordinate elements rendered`);
    
  } catch (error) {
    Logger.log(`❌ Error rendering coordinate elements: ${error.toString()}`);
  }
}

function renderCoordinateElement(slide, element, slideName, presentation) {
  if (!slide) {
    Logger.log(`❌ renderCoordinateElement called with undefined slide!`);
    return;
  }
  
  if (!element || !element.type) {
    Logger.log(`⚠️ Element or element.type is undefined`);
    return;
  }
  
  switch (element.type) {
    case 'text':
      renderCoordinateText(slide, element);
      break;
    
    // --- ADD THIS CASE ---
    case 'arcRectangle':
      renderCoordinateArcRectangle(slide, element);
      break;
      
    case 'rectangle':
      renderCoordinateRectangle(slide, element);
      break;
      
    case 'image':
      renderCoordinateImage(slide, element);
      break;
      
    case 'chart':
      renderCoordinateChart(slide, element, slideName);
      break;
      
    case 'table':
      renderCoordinateTable(slide, element, presentation);
      break;
      
    case 'shape':
      renderCoordinateShape(slide, element);
      break;
      
    case 'line':
      renderCoordinateLine(slide, element);
      break;
      
    case 'video':
      renderCoordinateVideo(slide, element);
      break;
      
    case 'arrayMap':
      renderCoordinateArrayMap(slide, element, slideName, presentation);
      break;
      
    case 'teamMember':
      renderCoordinateTeamMember(slide, element);
      break;
      
    default:
      Logger.log(`⚠️ Unknown element type: ${element.type}`);
  }
}

function renderCoordinateText(slide, element) {
  try {
    const textBox = slide.insertTextBox(element.content);
    
    textBox
      .setLeft(element.x)
      .setTop(element.y)
      .setWidth(element.width)
      .setHeight(element.height);
    
    //"Fill": "#FFC0CB","BACKGROUND": "#06065D", to apply background on the text.


    // 3. MEASURE THE BOX (Google Slides calculates the final height 'h' here)
    const x = textBox.getLeft();
    const y = textBox.getTop();
    const w = textBox.getWidth();
    const h = textBox.getHeight(); 

    // 4. BORDER CONFIGURATION
    const weight = element.borderWeight || 1.5;
    const color = element.borderColor || "#000000"; // Dynamic color per box

    // DRAW LEFT BORDER
    if (element.leftBorder) {
      const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x, y, x, y + h);
      line.getLineFill().setSolidFill(color);
      line.setWeight(weight);
      // Indent text so it doesn't touch the line
      textRange.getParagraphStyle().setIndentStart(10).setIndentFirstLine(10);
    }

    // DRAW RIGHT BORDER
    if (element.rightBorder) {
      const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x + w, y, x + w, y + h);
      line.getLineFill().setSolidFill(color);
      line.setWeight(weight);
    }

    // DRAW BOTTOM BORDER
    if (element.bottomBorder) {
      const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x, y + h, x + w, y + h);
      line.getLineFill().setSolidFill(color);
      line.setWeight(weight);
    }

    // 5. BACKGROUND & TRANSPARENCY
    // Use the keys from your JSON (BACKGROUND or backgroundColor)
    const bgColor = element.BACKGROUND || element.backgroundColor || element.Fill;
    if (bgColor) {
      textBox.getFill().setSolidFill(bgColor);
    } else {
      textBox.getFill().setTransparent();
    }
    
    
    // --- FONT LOGIC START ---
    // Define which font to use based on font size (e.g., if > 18, it's a heading)
    const headingThreshold = 20; 
    const isHeading = (element.fontSize >= headingThreshold);
    const selectedFont = isHeading ? 'Special Gothic Expanded One' : 'Figtree';
    // --- FONT LOGIC END ---

    let listItems = element.listItems;
    
    if (typeof listItems === 'string') {
      const trimmed = listItems.trim();
      if (trimmed.startsWith('[')) {
        try {
          listItems = JSON.parse(listItems);
        } catch (parseError) {
          listItems = [];
        }
      } else if (trimmed.startsWith('<')) {
        try {
          listItems = parseHtmlListItems(listItems);
        } catch (htmlError) {
          listItems = [];
        }
      }
    }
    
    if (element.isBulletList && Array.isArray(listItems)) {
      textBox.getText().clear();
      
      listItems.forEach((item, index) => {
        if (index > 0) {
          textBox.getText().appendText('\n');
        }
        textBox.getText().appendText('• ' + item);
      });
      
      const textStyle = textBox.getText().getTextStyle()
        .setFontSize(element.fontSize)
        .setFontFamily(selectedFont) // Applied dynamic font
        .setForegroundColor(element.color);
      
      if (element.fontWeight === 'bold') {
        textStyle.setBold(true);
      }
      
      const paragraphStyle = textBox.getText().getParagraphStyle();
      paragraphStyle
        .setSpaceAbove(0)
        .setSpaceBelow(element.lineSpacing || 6)
        .setIndentStart(0)
        .setIndentEnd(0)
        .setIndentFirstLine(0)
        .setLineSpacing(element.lineHeight || 1.15);
      
    } else {
      const textStyle = textBox.getText().getTextStyle()
        .setFontSize(element.fontSize)
        .setFontFamily(selectedFont) // Applied dynamic font
        .setForegroundColor(element.color);

      if (element.autoScale) {
        const content = (element.content || '').trim();
        const baseFontSize = element.fontSize;
        const boxWidth = element.width;
        const boxHeight = element.height;
        
        const charWidthRatio = element.fontWeight === 'bold' ? 0.45 : 0.42;
        const lineHeightRatio = 1.35;
        
        let bestFontSize = 0;
        
        for (let numLines = 1; numLines <= 5; numLines++) {
          const charsPerLine = Math.ceil(content.length / numLines);
          const fontSizeForWidth = boxWidth / (charsPerLine * charWidthRatio);
          const fontSizeForHeight = boxHeight / (numLines * lineHeightRatio);
          const candidateFontSize = Math.min(fontSizeForWidth, fontSizeForHeight);
          
          if (candidateFontSize <= baseFontSize && candidateFontSize > bestFontSize) {
            bestFontSize = candidateFontSize;
          }
        }
         
        const adjustedFontSize = Math.round(Math.min(baseFontSize, bestFontSize));
        textStyle.setFontSize(adjustedFontSize);
      }
      
      if (element.fontWeight === 'bold') {
        textStyle.setBold(true);
      }
      
      const paragraphStyle = textBox.getText().getParagraphStyle();
      paragraphStyle
        .setSpaceAbove(0)
        .setSpaceBelow(0)
        .setIndentStart(0)
        .setIndentEnd(0)
        .setIndentFirstLine(0);
      
      if (element.lineHeight) {
        paragraphStyle.setLineSpacing(element.lineHeight);
      } else {
        paragraphStyle.setLineSpacing(1.0);
      }
      
      if (element.textAlign) {
        switch (element.textAlign.toLowerCase()) {
          case 'center':
            paragraphStyle.setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
            break;
          case 'right':
            paragraphStyle.setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
            break;
          case 'justify':
            paragraphStyle.setParagraphAlignment(SlidesApp.ParagraphAlignment.JUSTIFIED);
            break;
          case 'left':
          default:
            paragraphStyle.setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
            break;
        }
      }
    }
    
  } catch (error) {
    Logger.log(`⚠️ Error rendering text element: ${error.toString()}`);
  }
}

// Helper function to parse HTML list items
function parseHtmlListItems(htmlString) {
  if (!htmlString) return [];
  
  const items = [];
  
  // Extract content between <li> tags using regex
  const liRegex = /<li[^>]*>(.*?)<\/li>/gs;
  let match;
  
  while ((match = liRegex.exec(htmlString)) !== null) {
    // Remove all HTML tags and get plain text
    let text = match[1]
      .replace(/<[^>]+>/g, '') // Remove all HTML tags
      .replace(/&nbsp;/g, ' ') // Replace &nbsp; with space
      .replace(/&amp;/g, '&')  // Replace &amp; with &
      .replace(/&lt;/g, '<')   // Replace &lt; with 
      .replace(/&gt;/g, '>')   // Replace &gt; with >
      .trim();
    
    if (text) {
      items.push(text);
    }
  }
  
  return items;
}

function renderCoordinateArcRectangle(slide, element) {
  try {
    Logger.log(`🟦 Rendering Arc Rectangle at (${element.x}, ${element.y})`);
    
    // 1. Insert the Rounded Rectangle shape
    const shape = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE);
    
    // 2. Set dimensions and position
    shape.setLeft(element.x)
         .setTop(element.y)
         .setWidth(element.width)
         .setHeight(element.height);
    
    // 3. Set Background Fill Color
    if (element.fill) {
      shape.getFill().setSolidFill(element.fill);
    } else {
      shape.getFill().setTransparent();
    }

    // 4. Set Border (optional)
    if (element.borderColor) {
      shape.getBorder().getLineFill().setSolidFill(element.borderColor);
      shape.getBorder().setWeight(element.borderWeight || 1);
    } else {
      shape.getBorder().setTransparent();
    }

    // 5. SET THE ARC RADIUS (The curvature)
    // The value ranges from 0.0 to 0.5 (0.5 is a full pill shape)
    if (element.arcRadius !== undefined) {
      const adjustments = shape.getAdjustments();
      if (adjustments.length > 0) {
        adjustments[0].setValue(element.arcRadius);
      }
    }
    
    Logger.log(`✅ Arc Rectangle rendered successfully`);
    
  } catch (error) {
    Logger.log(`⚠️ Error rendering arc rectangle: ${error.toString()}`);
  }
}

function renderCoordinateRectangle(slide, element) {
  try {
    const rect = slide.insertShape(SlidesApp.ShapeType.RECTANGLE);
    
    rect
      .setLeft(element.x)
      .setTop(element.y)
      .setWidth(element.width)
      .setHeight(element.height);
    
    rect.getFill().setSolidFill(element.fill);
    rect.getBorder().setTransparent();
    
  } catch (error) {
    Logger.log(`⚠️ Error rendering rectangle element: ${error.toString()}`);
  }
}

function renderCoordinateImage(slide, element) {
  try {
    Logger.log(`🖼️ Rendering image from URL: ${element.url}`);
    
    const image = slide.insertImage(element.url);
    
    image
      .setLeft(element.x)
      .setTop(element.y)
      .setWidth(element.width)
      .setHeight(element.height);
    
    Logger.log(`✅ Image rendered at (${element.x}, ${element.y}) with size ${element.width}x${element.height}`); 
    
  } catch (error) {
    Logger.log(`⚠️ Error rendering image element: ${error.toString()}`);
  }
}

// Enhanced renderCoordinateChart function with pie chart label styling
function renderCoordinateChart(slide, element, slideName) {
  try {
    if (!slide || !element) {
      Logger.log(`❌ Error: Invalid parameters for chart rendering`);
      return;
    }
    
    // Enhanced data handling - support both object and JSON string
    let chartData = element.data;
    
    // If data is a string, try to parse it as JSON
    if (typeof chartData === 'string' && chartData.trim().startsWith('{')) {
      try {
        chartData = JSON.parse(chartData);
        Logger.log(`📊 Parsed chart data from JSON string`);
      } catch (parseError) {
        Logger.log(`⚠️ Failed to parse chart data JSON: ${parseError.toString()}`);
        return;
      }
    }
    
    // Ensure we have valid chart data
    if (!chartData || !chartData.columns || !chartData.rows) {
      Logger.log(`❌ Error: Invalid chart data structure`);
      return;
    }
    
    // Increment chart counter
    chartCounter++;
    
    Logger.log(`📊 Rendering chart #${chartCounter} for ${slideName}`);
    Logger.log(`   Chart data: ${chartData.rows.length} rows, ${chartData.columns.length} columns`);
    
    // Open the spreadsheet
    const spreadsheet = SpreadsheetApp.openById(SHEETS_CONFIG.SPREADSHEET_ID);
    
    // Create sheet name
    const sheetName = `${slideName}_Chart${chartCounter}`;
    const sheet = spreadsheet.insertSheet(sheetName);
    
    // Prepare data from JSON
    const headers = chartData.columns.map(col => col.label);
    const dataRows = chartData.rows;
    const tableData = [headers, ...dataRows];
    
    // Write data to sheet
    const numRows = tableData.length;
    const numCols = headers.length;
    const range = sheet.getRange(1, 1, numRows, numCols);
    range.setValues(tableData);
    
    // Format header row (matching reference)
    const headerRange = sheet.getRange(1, 1, 1, numCols);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(12);
    headerRange.setHorizontalAlignment('center');
    
    // Format data rows
    if (numRows > 1) {
      const dataRange = sheet.getRange(2, 1, numRows - 1, numCols);
      dataRange.setFontSize(11);
      
      for (let i = 2; i <= numRows; i++) {
        const rowRange = sheet.getRange(i, 1, 1, numCols);
        if (i % 2 === 0) {
          rowRange.setBackground('#f8f9fa');
        } else {
          rowRange.setBackground('#ffffff');
        }
      }
    }
    
    // Add borders
    range.setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
    sheet.autoResizeColumns(1, numCols);
    
    // Create data range string
    const dataRangeString = `A2:${String.fromCharCode(64 + numCols)}${numRows}`;
    
    Logger.log(`   Data range: ${dataRangeString}`);
    
    // Clear any existing charts first
    const existingCharts = sheet.getCharts();
    existingCharts.forEach(chart => sheet.removeChart(chart));
    
    // Get the data range object
    const dataRangeObject = sheet.getRange(dataRangeString);
    
    // Determine chart type and set appropriate dimensions
    let chartType = Charts.ChartType.COLUMN;
    let chartWidth = 750;  // Default for column chart
    let chartHeight = 400; // Default for column chart
    
    if (element.chartType) {
      switch (element.chartType.toLowerCase()) {
        case 'line':
          chartType = Charts.ChartType.LINE;
          chartWidth = 750;
          chartHeight = 400;
          break;
        case 'bar':
          chartType = Charts.ChartType.BAR;
          chartWidth = 400;   // Bar chart specific
          chartHeight = 1100; // Bar chart specific
          break;
        case 'area':
          chartType = Charts.ChartType.AREA;
          chartWidth = 750;
          chartHeight = 400;
          break;
        case 'pie':
          chartType = Charts.ChartType.PIE;
          chartWidth = 400;   // Pie chart specific
          chartHeight = 400;  // Pie chart specific
          break;
        default:
          chartType = Charts.ChartType.COLUMN;
          chartWidth = 800;   // Column chart specific
          chartHeight = 400;  // Column chart specific
      }
    }
    
    Logger.log(`📊 Chart type: ${element.chartType}, Dimensions: ${chartWidth}x${chartHeight}`);
    
    // Better color support for pie charts
    let chartColors = element.colors || ['#4285f4'];
    if (element.chartType && element.chartType.toLowerCase() === 'pie' && (!element.colors || element.colors.length === 1)) {
      // Use blue palette for pie charts if only one color or no colors specified
      chartColors = ['#4285F4', '#87CEEB', '#5F9EA0', '#B0E0E6', '#ADD8E6'];
    }
    
    // Create chart builder with chart-type-specific dimensions
    let chartBuilder = sheet.newChart()
      .setChartType(chartType)
      .addRange(dataRangeObject)
      .setPosition(15, 1, 0, 0)
      .setOption('title', element.title || '')
      .setOption('titleTextStyle.fontSize', 16)
      .setOption('titleTextStyle.bold', true)
      .setOption('width', chartWidth)                     // ← Chart-type-specific width
      .setOption('height', chartHeight)                   // ← Chart-type-specific height
      .setOption('hAxis.title', element.xAxisTitle || '')
      .setOption('vAxis.title', element.yAxisTitle || '')
      .setOption('backgroundColor', '#E7F2F4')
      .setOption('vAxis.format', '#,###')
      .setOption('colors', chartColors);
    
    // Add pie chart specific styling
    if (element.chartType && element.chartType.toLowerCase() === 'pie') {
      Logger.log(`📊 Applying pie chart label styling (fontSize: 15)`);
      
      // Style for slice labels (text on the pie slices)
      chartBuilder = chartBuilder
        .setOption('pieSliceTextStyle.fontSize', 15)
        .setOption('pieSliceTextStyle.color', '#333333')
        .setOption('pieSliceTextStyle.bold', true);
      
      // Style for legend labels
      chartBuilder = chartBuilder
        .setOption('legend.textStyle.fontSize', 15)
        .setOption('legend.textStyle.color', '#333333')
        .setOption('legend.position', 'bottom')
        .setOption('legend.alignment', 'center');
    }
    
    const chart = chartBuilder.build();
    
    sheet.insertChart(chart);
    
    Logger.log(`📊 Chart created for sheet: ${sheetName}`);
    Utilities.sleep(4000);
    
    // Get the chart as image
    const charts = sheet.getCharts();
    if (charts.length === 0) {
      throw new Error('No chart found after creation');
    }
    
    const chartBlob = charts[0].getAs('image/png');
    const chartImage = slide.insertImage(chartBlob);
    
    chartImage
      .setLeft(element.x)
      .setTop(element.y)
      .setWidth(element.width)
      .setHeight(element.height);
    
    Logger.log(`✅ Chart ${chartCounter} rendered successfully`);
    
  } catch (error) {
    Logger.log(`❌ Failed to create chart: ${error.toString()}`);
    
    // Create placeholder
    if (slide) {
      const bg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE);
      bg.setLeft(element.x)
        .setTop(element.y)
        .setWidth(element.width)
        .setHeight(element.height);
      bg.getFill().setSolidFill('#ffeeee');
      bg.getBorder().getLineFill().setSolidFill('#ff0000');
    }
  }
}

function renderCoordinateTable(slide, element, presentation) {
  try {
    Logger.log(`📋 Rendering table...`);
    
    let tableData = element.data;
    let dataArray;
    let rows, columns;
    
    // Parse data
    if (typeof tableData === 'string' && (tableData.includes('###') || tableData.includes('|||'))) {
      const rowStrings = tableData.split('###');
      dataArray = rowStrings.map(rowStr => rowStr.split('|||'));
      rows = dataArray.length;
      columns = dataArray[0] ? dataArray[0].length : 0;
    }
    else if (typeof tableData === 'string' && tableData.trim().startsWith('[')) {
      tableData = JSON.parse(tableData);
      if (Array.isArray(tableData) && tableData.length > 0 && typeof tableData[0] === 'object') {
        const columnKeys = ['Key Workstreams', 'Customer Owner', 'Docebo Owner', 'Priority', 'Status', 'Due Date'];
        dataArray = tableData.map(row => columnKeys.map(key => row[key] || ''));
        // Add column headers as the first row
        dataArray.unshift(columnKeys);
        rows = dataArray.length;
        columns = columnKeys.length;
      }
    }
    else if (Array.isArray(tableData)) {
      // Check if it's an array of objects (not arrays)
      if (tableData.length > 0 && typeof tableData[0] === 'object' && !Array.isArray(tableData[0])) {
        // Convert array of objects to array of arrays
        const columnKeys = ['Key Workstreams', 'Customer Owner', 'Docebo Owner', 'Priority', 'Status', 'Due Date'];
        dataArray = tableData.map(row => columnKeys.map(key => row[key] || ''));
        // Add column headers as the first row
        dataArray.unshift(columnKeys);
        rows = dataArray.length;
        columns = columnKeys.length;
      } else {
        // It's an array of arrays
        dataArray = tableData;
        rows = element.rows || dataArray.length;
        columns = element.columns || (dataArray[0] ? dataArray[0].length : 0);
      }
    }
    else {
      Logger.log(`⚠️ No valid table data provided`);
      return;
    }
    
    Logger.log(`📋 Creating table with ${rows} rows and ${columns} columns`);
    
    // Create table using Apps Script API
    const table = slide.insertTable(rows, columns);
    
    // 🔍 DEBUG: Log element.width
    Logger.log(`🔍 DEBUG: element.width = ${element.width}`);
    Logger.log(`🔍 DEBUG: element.width type = ${typeof element.width}`);
    
    // Set position
    table.setLeft(element.x).setTop(element.y);
    
    // Try to set width with error handling
    if (element.width !== undefined && element.width !== null) {
      try {
        Logger.log(`🔍 DEBUG: Attempting to set width to ${element.width}`);
        table.setWidth(element.width);
        Logger.log(`✅ Width set successfully`);
      } catch (widthError) {
        Logger.log(`⚠️ Error setting table width: ${widthError.toString()}`);
      }
    } else {
      Logger.log(`⚠️ element.width is undefined or null, skipping setWidth`);
    }
    
    Logger.log(`🔍 DEBUG: About to populate cells...`);
    

    // Populate and style cells
    dataArray.forEach((row, rowIndex) => {
      row.forEach((cellContent, colIndex) => {
        if (rowIndex < rows && colIndex < columns) {
          const cell = table.getCell(rowIndex, colIndex);
          cell.getText().setText(cellContent || '');
          
          const textStyle = cell.getText().getTextStyle();
          const paragraphStyle = cell.getText().getParagraphStyle();
          
          // Header row styling
          if (rowIndex === 0) {
            cell.getFill().setSolidFill(element.headerStyle?.backgroundColor || '#EFF4FF');
            textStyle
              .setFontSize(element.headerStyle?.fontSize || 11)
              .setFontFamily(element.headerStyle?.fontFamily || 'Figtree')
              .setForegroundColor(element.headerStyle?.color || '#101828')
              .setBold(true);
          }
          // Data row styling
          else {
            cell.getFill().setSolidFill(element.cellBackground || '#FFFFFF');
            textStyle
              .setFontSize(element.cellStyle?.fontSize || 11)
              .setFontFamily(element.cellStyle?.fontFamily || 'Figtree')
              .setForegroundColor(element.cellStyle?.color || '#344054');
            
            // Column-specific styling
            if (element.columnStyles && element.columnStyles[colIndex]) {
              const colStyle = element.columnStyles[colIndex];
              if (colStyle.color) textStyle.setForegroundColor(colStyle.color);
              if (colStyle.fontWeight === 'bold') textStyle.setBold(true);
              if (colStyle.backgroundColor) cell.getFill().setSolidFill(colStyle.backgroundColor);
            }
          }
          
          paragraphStyle.setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
        }
      });
    });
    
    Logger.log(`✅ Table rendered with ${rows} rows and ${columns} columns`);
    
  } catch (error) {
    Logger.log(`⚠️ Error rendering table element: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
  }
}

// Helper function to convert hex color to RGB object
function hexToRgb(hex) {
  const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result ? {
    red: parseInt(result[1], 16) / 255,
    green: parseInt(result[2], 16) / 255,
    blue: parseInt(result[3], 16) / 255
  } : { red: 0, green: 0, blue: 0 };
}

function renderCoordinateShape(slide, element) {
  try {
    Logger.log(`🔷 Rendering shape: ${element.shapeType}`);
    
    let shapeType;
    switch (element.shapeType) {
      case 'circle':
        shapeType = SlidesApp.ShapeType.ELLIPSE;
        break;
      case 'triangle':
        shapeType = SlidesApp.ShapeType.TRIANGLE;
        break;
      case 'diamond':
        shapeType = SlidesApp.ShapeType.DIAMOND;
        break;
      case 'star':
        shapeType = SlidesApp.ShapeType.STAR_5;
        break;
      case 'arrow':
        shapeType = SlidesApp.ShapeType.RIGHT_ARROW;
        break;
      case 'roundRectangle':
        shapeType = SlidesApp.ShapeType.ROUND_RECTANGLE;
        break;
      
      case 'arcRectangle':
        shapeType = SlidesApp.ShapeType.ARC_RECTANGLE;
        break;

      default:
        shapeType = SlidesApp.ShapeType.RECTANGLE;
    }
    
    const shape = slide.insertShape(shapeType);

    shape
      .setLeft(element.x)
      .setTop(element.y)
      .setWidth(element.width)
      .setHeight(element.height);
    
    if (element.fill) {
      shape.getFill().setSolidFill(element.fill);
    }

    if (element.borderColor) {
      const border = shape.getBorder();
      border.getLineFill().setSolidFill(element.borderColor);
      border.setWeight(element.borderWeight || 1);
    } else {
      shape.getBorder().setTransparent();
    }
    
    Logger.log(`✅ Shape rendered`);
    
  } catch (error) {
    Logger.log(`⚠️ Error rendering shape element: ${error.toString()}`);
  }
}

function renderCoordinateLine(slide, element) {
  try {
    Logger.log(`📏 Rendering line`);
    
    const line = slide.insertLine(
      SlidesApp.LineCategory.STRAIGHT,
      element.startX,
      element.startY,
      element.endX,
      element.endY
    );
    
    line.getLineFill().setSolidFill(element.color || '#000000');
    line.setWeight(element.weight || 1);
    
    if (element.dashStyle === 'dashed') {
      line.setDashStyle(SlidesApp.DashStyle.DASH);
    } else if (element.dashStyle === 'dotted') {
      line.setDashStyle(SlidesApp.DashStyle.DOT);
    }
    
    Logger.log(`✅ Line rendered`);
    
  } catch (error) {
    Logger.log(`⚠️ Error rendering line element: ${error.toString()}`);
  }
}

function renderCoordinateVideo(slide, element) {
  try {
    Logger.log(`🎥 Rendering video from URL: ${element.url}`);
    
    const video = slide.insertVideo(element.url);
    
    video
      .setLeft(element.x)
      .setTop(element.y)
      .setWidth(element.width)
      .setHeight(element.height);
    
    Logger.log(`✅ Video rendered`);
    
  } catch (error) {
    Logger.log(`⚠️ Error rendering video element: ${error.toString()}`);
  }
}

function renderCoordinateArrayMap(slide, element, slideName, presentation) {
  try {
    // Enhanced items handling - support both array and JSON string
    let items = element.items;
    
    // If items is a string, try to parse it as JSON
    if (typeof items === 'string' && items.trim().startsWith('[')) {
      try {
        items = JSON.parse(items);
        Logger.log(`🔄 Parsed items from JSON string: ${items.length} items`);
      } catch (parseError) {
        Logger.log(`⚠️ Failed to parse items JSON: ${parseError.toString()}`);
        items = []; // Fallback to empty array
      }
    }
    
    // Ensure items is an array
    if (!Array.isArray(items)) {
      Logger.log(`⚠️ Items is not an array, skipping arrayMap render`);
      return;
    }
    
    Logger.log(`🔄 Rendering array map with ${items.length} items`);
    
    items.forEach((item, index) => {
      let x = element.x;
      let y = element.y;
      
      if (element.layout === 'vertical') {
        y = element.y + (index * element.spacing);
      } else if (element.layout === 'horizontal') {
        x = element.x + (index * element.spacing);
      } else if (element.layout === 'grid' && element.columns) {
        const row = Math.floor(index / element.columns);
        const col = index % element.columns;
        x = element.x + (col * element.spacing);
        y = element.y + (row * (element.rowSpacing || element.spacing));
      }
      
      const mappedElement = {
        ...element.template,
        x: x,
        y: y
      };
      
      if (mappedElement.type === 'text' && mappedElement.content) {
        mappedElement.content = mappedElement.content.replace('{{item}}', item);
        mappedElement.content = mappedElement.content.replace('{{index}}', index + 1);
      }
      
      if (element.template.dynamicProps) {
        Object.keys(element.template.dynamicProps).forEach(prop => {
          if (typeof item === 'object' && item[element.template.dynamicProps[prop]]) {
            mappedElement[prop] = item[element.template.dynamicProps[prop]];
          }
        });
      }
      
      renderCoordinateElement(slide, mappedElement, slideName, presentation);
    });
    
    Logger.log(`✅ Array map rendered with ${items.length} items`);
    
  } catch (error) {
    Logger.log(`⚠️ Error rendering array map: ${error.toString()}`);
  }
}

function renderCoordinateTeamMember(slide, element) {
  try {
    Logger.log(`👤 Rendering team member card: ${element.name}`);
    
    // Profile picture (rounded rectangle on the left)
    let profileImageUrl = element.profilePicture;
    if (!profileImageUrl || profileImageUrl === 'Unavailable' || profileImageUrl.trim() === '') {
      profileImageUrl = PLACEHOLDER_IMAGE_URL;
      Logger.log(`   Using placeholder image for ${element.name}`);
    }
    
    try {
      // Insert profile image as rounded rectangle
      const profileImage = slide.insertImage(profileImageUrl);
      profileImage
        .setLeft(element.x + 11) // 8px padding from left edge of card
        .setTop(element.y + 6)  // 8px padding from top edge of card
        .setWidth(74)           // Square image
        .setHeight(71);
      
      Logger.log(`   Profile picture inserted for ${element.name}`);
    } catch (imageError) {
      Logger.log(`   ⚠️ Image insertion failed, using placeholder rectangle: ${imageError.toString()}`);
      
      // Create placeholder rounded rectangle if image fails
      const placeholderRect = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE);
      placeholderRect
        .setLeft(element.x + 13)
        .setTop(element.y + 5)
        .setWidth(74)
        .setHeight(74);
      placeholderRect.getFill().setSolidFill('#e0e0e0');
      placeholderRect.getBorder().setTransparent();
      
      // Add placeholder text
      const placeholderText = slide.insertTextBox('Profile');
      placeholderText
        .setLeft(element.x + 8)
        .setTop(element.y + 28)
        .setWidth(64)
        .setHeight(24);
      placeholderText.getFill().setTransparent();
      placeholderText.getBorder().setTransparent();
      placeholderText.getText().getTextStyle()
        .setFontSize(10)
        .setForegroundColor('#666666');
      placeholderText.getText().getParagraphStyle()
        .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    }
    
    // Clean HTML tags from name
    const cleanName = stripHtmlTags(element.name);
    
    // Name text (bold, positioned to the right of the image)
    const nameText = slide.insertTextBox(cleanName);
    nameText
      .setLeft(element.x + 101) // Start after image (8 + 64 + 8 spacing)
      .setTop(element.y + 19)  // Vertically centered in upper part of card -2
      .setWidth(element.width - 50) // Remaining width of card minus padding
      .setHeight(16);
    
    nameText.getFill().setTransparent();
    nameText.getBorder().setTransparent();
    
    const nameStyle = nameText.getText().getTextStyle()
      .setFontSize(12)
      .setFontFamily('Figtree')
      .setForegroundColor('#000000')
      .setBold(true);
    
    const nameParagraph = nameText.getText().getParagraphStyle();
    nameParagraph.setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
    
    // Membership/Role text (smaller font, positioned below name)
    const membershipText = slide.insertTextBox(element.membership);
    membershipText
      .setLeft(element.x + 101) // Same x as name
      .setTop(element.y + 35)  // Below the name
      .setWidth(element.width - 50) // Same width as name
      .setHeight(20);
    
    membershipText.getFill().setTransparent();
    membershipText.getBorder().setTransparent();
    
    const membershipStyle = membershipText.getText().getTextStyle()
      .setFontSize(12)
      .setFontFamily('Figtree')
      .setForegroundColor('#666666');
    
    const membershipParagraph = membershipText.getText().getParagraphStyle();
    membershipParagraph.setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
    
    Logger.log(`✅ Team member card rendered: ${cleanName}`);
    
  } catch (error) {
    Logger.log(`⚠️ Error rendering team member card: ${error.toString()}`);
  }
}

function cleanSpreadsheetOnly() {
  try {
    Logger.log('🧹 RUNNING STANDALONE SPREADSHEET CLEANUP...');
    
    const spreadsheet = SpreadsheetApp.openById(SHEETS_CONFIG.SPREADSHEET_ID);
    const originalSheetCount = spreadsheet.getSheets().length;
    
    
    Logger.log(`✅ Cleanup complete! Deleted ${originalSheetCount - 1} sheets`);
    Logger.log(`📊 Spreadsheet URL: ${spreadsheet.getUrl()}`);
    
    return {
      success: true,
      sheetsDeleted: originalSheetCount - 1,
      url: spreadsheet.getUrl()
    };
    
  } catch (error) {
    Logger.log(`❌ Cleanup failed: ${error.toString()}`);
    return {
      success: false,
      error: error.toString()
    };
  }
}

function testCoordinateSlides() {
  const result = createCoordinateSlides();
  
  if (result.success) {
    console.log(`🎯 SUCCESS! Coordinate slides: ${result.url}`);
    console.log(`📊 Slides created: ${result.slidesCreated}`);
    console.log(`📈 Charts created: ${result.chartsCreated}`);
    Logger.log(`🎉 COORDINATE SLIDES TEST COMPLETED!`);
    Logger.log(`🔗 Presentation URL: ${result.url}`);
    Logger.log(`📊 Charts saved in spreadsheet: https://docs.google.com/spreadsheets/d/${SHEETS_CONFIG.SPREADSHEET_ID}`);
    return result.url;
  } else {
    console.log('❌ FAILED! Check logs for errors.');
    console.log(`Error: ${result.error}`);
    Logger.log(`💥 COORDINATE SLIDES TEST FAILED: ${result.error}`);
    return null;
  }
}

function verifySpreadsheetAccess() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEETS_CONFIG.SPREADSHEET_ID);
    Logger.log(`✅ Successfully accessed spreadsheet: ${spreadsheet.getName()}`);
    Logger.log(`   Spreadsheet URL: ${spreadsheet.getUrl()}`);
    
    // List existing sheets
    const sheets = spreadsheet.getSheets();
    Logger.log(`   Existing sheets: ${sheets.map(s => s.getName()).join(', ')}`);
    
    return true;
  } catch (error) {
    Logger.log(`❌ Cannot access spreadsheet: ${error.toString()}`);
    Logger.log(`   Make sure you have access to: https://docs.google.com/spreadsheets/d/${SHEETS_CONFIG.SPREADSHEET_ID}`);
    return false;
  }
}

function listAllSheets() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SHEETS_CONFIG.SPREADSHEET_ID);
    const sheets = spreadsheet.getSheets();
    
    Logger.log('📋 CURRENT SHEETS IN SPREADSHEET:');
    Logger.log('═══════════════════════════════════════');
    sheets.forEach((sheet, index) => {
      const charts = sheet.getCharts();
      Logger.log(`${index + 1}. "${sheet.getName()}" - ${charts.length} chart(s)`);
    });
    Logger.log('═══════════════════════════════════════');
    Logger.log(`Total: ${sheets.length} sheet(s)`);
    
    return sheets.map(s => ({
      name: s.getName(),
      charts: s.getCharts().length
    }));
    
  } catch (error) {
    Logger.log(`❌ Error listing sheets: ${error.toString()}`);
    return [];
  }
}

function testSimpleChart() {
  try {
    Logger.log('🧪 Testing simple chart creation...');
    
    const spreadsheet = SpreadsheetApp.openById(SHEETS_CONFIG.SPREADSHEET_ID);
    const testSheet = spreadsheet.insertSheet('Test_Chart_' + new Date().getTime());
    
    // Simple data
    const data = [
      ['Month', 'Active Users'],
      ['May 2025', 2078],
      ['April 2025', 1134],
      ['March 2025', 1115]
    ];
    
    // Write data
    testSheet.getRange(1, 1, 4, 2).setValues(data);
    
    // Format header
    testSheet.getRange(1, 1, 1, 2)
      .setBackground('#4285f4')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    
    // Create chart exactly as in reference
    const dataRange = testSheet.getRange('A1:B4');
    
    const chart = testSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dataRange)
      .setPosition(15, 1, 0, 0)
      .setOption('title', 'Monthly Active Users')
      .setOption('titleTextStyle.fontSize', 16)
      .setOption('titleTextStyle.bold', true)
      .setOption('width', 600)
      .setOption('height', 400)
      .setOption('hAxis.title', 'Month')
      .setOption('vAxis.title', 'Active Users')
      .setOption('backgroundColor', '#ffffff')
      .setOption('vAxis.format', '#,###')
      .setOption('colors', ['#4285f4'])
      .build();
    
    testSheet.insertChart(chart);
    
    Logger.log(`✅ Chart created`);
    Logger.log(`📊 View sheet: ${testSheet.getUrl()}`);
    
    Utilities.sleep(4000);
    
    const charts = testSheet.getCharts();
    Logger.log(`Charts found: ${charts.length}`);
    
  } catch (error) {
    Logger.log(`❌ Test error: ${error.toString()}`);
  }
}

function sharePresentationWithSpecificEmails(presentation) {
  let emailList = ["nitish.kumar@unifyapps.com"]
  try {
    Logger.log(`📧 Sharing presentation to specific emails...`);
    
    Utilities.sleep(2000);
    
    const file = DriveApp.getFileById(presentation.getId());
    Logger.log(`📁 File: ${file.getName()}`);
    
    // Share with specific email addresses
    emailList.forEach(email => {
      try {
        file.addEditor(email);
        Logger.log(`✅ Added editor: ${email}`);
      } catch (emailError) {
        Logger.log(`❌ Failed to add ${email}: ${emailError.toString()}`);
      }
    });
    
    Logger.log(`📧 Sharing complete for ${emailList.length} recipients`);
    Logger.log(`🔗 Presentation URL: ${presentation.getUrl()}`);
    
  } catch (error) {
    Logger.log(`❌ Email sharing error: ${error.toString()}`);
  }
}
