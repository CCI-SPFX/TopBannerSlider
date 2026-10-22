import * as React from "react";
import { useState, useEffect } from "react";
// import DOMPurify from "dompurify";
// import styles from './Lab02.module.scss';
import type { ILab02Props } from "./ILab02Props";
import "./style.css";
import bannerBGImg from "../assets/Bg.png";
// import { escape } from '@microsoft/sp-lodash-subset';

import { Swiper, SwiperSlide } from "swiper/react";

// import Swiper and modules styles
import "swiper/css";
import "swiper/css/navigation";
import "swiper/css/pagination";
import { Autoplay, Pagination } from "swiper";

// const sanitize = (html: string) =>
//   DOMPurify.sanitize(html ?? "", {
//   ALLOWED_TAGS: ["b","strong","i","em","u","br","p","div","span","a","ul","ol","li","h1","h2","h3"],
//   ALLOWED_ATTR: ["href","target","rel","style","class"],
// });

const Lab02: React.FC<ILab02Props> = ({
  slider,
  listName,
  siteUrl,
  timer,
  showContent,

  contentAlignment,
  verticalAlign,
  heading,
  headingSize,
  headingColor,
  headingShow,

  description,
  descriptionSize,
  descriptionColor,
  descriptionShow,

  btnText,
  btnTextSize,
  btnBgColor,
  btnBorderRadius,
  btnLink,
  BtnTextColor,
  btnShow,

  imgUrl,
  bgType,
  itemHeight,
  aspectRatio,
}) => {
  const [heroSlider, setHeroSlider] = useState<boolean>(false);
  const [heroBanner, setHeroBanner] = useState<boolean>(true);

  const [items, setItems] = useState<any[]>([]);
  const apiUrl = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Title,Description,ButtonLink,isActive&$filter=isActive eq 1&$expand=AttachmentFiles`;

  const bgImage = imgUrl?.trim() ? imgUrl : bannerBGImg;

  useEffect(() => {
    if (slider === true) {
      setHeroSlider(true);
      setHeroBanner(false);
    } else {
      setHeroBanner(true);
      setHeroSlider(false);
    }
  }, [slider]);

  useEffect(() => {
    if (btnBgColor) {
      document.documentElement.style.setProperty(
        "--swiper-pagination-color",
        btnBgColor,
      );
    }
  }, [btnBgColor]);

  useEffect(() => {
    const fetchData = async () => {
      try {
        const response = await fetch(apiUrl, {
          method: "Get",
          headers: {
            Accept: "application/json;odata=verbose",
          },
        });
        const data: any = await response.json();

        // Process items to add the full image URL if available
        const itemsWithImages = data.d.results.map((item: any) => {
          const imageUrl =
            item.AttachmentFiles.results.length > 0 ?
              `${item.AttachmentFiles.results[0].ServerRelativeUrl}`
            : null;
          return { ...item, imageUrl };
        });
        setItems(itemsWithImages);
      } catch (error) {
        console.log("Error fetching data:", error);
      }
    };
    fetchData();
  }, [siteUrl, listName]);

  return (
    <>
      {heroBanner && (
        <section
          className={`topBanner ${
            contentAlignment === "Right" ? "contRight"
            : contentAlignment === "Center" ? "contCenter"
            : "contLeft"
          }`}
          style={
            showContent ?
              { minHeight: itemHeight !== "" ? itemHeight : "400px" }
            : { aspectRatio: aspectRatio !== "" ? aspectRatio : "22/6" }
          }
        >
          {!showContent ?
            <a
              href={btnLink}
              className="imgDiv"
              data-overlay={bgType}
              target="_blank"
              rel="noopener noreferrer"
              data-interception="off"
            >
              <img src={bgImage} alt="BG image" />
            </a>
          : <div className="imgDiv" data-overlay={bgType}>
              <img src={bgImage} alt="BG image" {...{ fetchPriority: "high" }} loading="eager"/>
            </div>
          }
          {showContent && (
            <div
              className={`wcContent ${
                verticalAlign === "Top" ? "vertTop"
                : verticalAlign === "Bottom" ? "vertBottom"
                : "Center"
              }`}
              style={
            showContent ?
              { minHeight: itemHeight !== "" ? itemHeight : "400px" }
            :{}
              }
            >
              <h2
                className="wcNote"
                style={{
                  fontSize: headingSize !== "" ? headingSize : "42px",
                  color: headingColor !== "" ? headingColor : "#fff",
                  display: headingShow === true ? "-webkit-box" : "none",
                }}
              >
                {heading !== "" ? heading : "Heading Text"}
              </h2>

              <p
                style={{
                  fontSize: descriptionSize !== "" ? descriptionSize : "18px",
                  color: descriptionColor !== "" ? descriptionColor : "#fff",
                  display: descriptionShow === true ? "-webkit-box" : "none",
                }}
              >
                {description !== "" ? description : "Description Text"}
              </p>

              <a
                href={btnLink}
                target="_blank"
                rel="noopener noreferrer"
                data-interception="off"
                style={{
                  fontSize: btnTextSize !== "" ? btnTextSize : "#fff",
                  backgroundColor: btnBgColor !== "" ? btnBgColor : "#fff",
                  borderRadius:
                    btnBorderRadius !== "" ? btnBorderRadius : "0px",
                  color: BtnTextColor !== "" ? BtnTextColor : "#fff",
                  display: btnShow === true ? "block" : "none",
                }}
              >
                {btnText !== "" ? btnText : "Button Text"}
              </a>
            </div>
          )}
        </section>
        // </a>
      )}

      {heroSlider && (
        <Swiper
          spaceBetween={10}
          loop={true}
          speed={1200}
          autoplay={{
            delay: timer * 1000,
            disableOnInteraction: false,
            pauseOnMouseEnter: true,
          }}
          pagination={{
            clickable: true,
          }}
          modules={[Autoplay, Pagination]}
          className="mySwiper"
        >
          {items.map((item, index) => (
            <SwiperSlide
              key={index}
              style={
                !showContent ?
                  { aspectRatio: aspectRatio !== "" ? aspectRatio : "22/6" }
                : {}
              }
            >
              {/* <a
                href={item.ButtonLink?.Url || "#"}
                target="_blank"
                rel="noopener noreferrer"
                data-interception="off"
                style={
                  showContent ?
                    { minHeight: itemHeight !== "" ? itemHeight : "400px" }
                  : {}
                }
              > */}
                <div
                  className={`topBanner ${
                    contentAlignment === "Right" ? "contRight"
                    : contentAlignment === "Center" ? "contCenter"
                    : "contLeft"
                  }`}
                  // style={{minHeight:itemHeight !== "" ? itemHeight : "100%"}}
                >



{!showContent ?
            <a
              href={item.ButtonLink?.Url || "#"}
              className="imgDiv"
              data-overlay={bgType}
              target="_blank"
              rel="noopener noreferrer"
              data-interception="off"
            >
              <img src={item.imageUrl || bgImage} alt="BG image" />
            </a>
          : <div className="imgDiv" data-overlay={bgType}>
              <img src={item.imageUrl || bgImage} alt="BG image" {...{ fetchPriority: "high" }} loading="eager"/>
            </div>
          }


                  {/* <div className="imgDiv" data-overlay={bgType}>
                    <img src={item.imageUrl || bgImage} alt="BG image" {...{ fetchPriority: "high" }} loading="eager" />
                  </div> */}
                  {showContent && (
                    <div
                      className={`wcContent ${
                        verticalAlign === "Top" ? "vertTop"
                        : verticalAlign === "Bottom" ? "vertBottom"
                        : "Center"
                      }`} 
                      style={
                  showContent ?
                    { minHeight: itemHeight !== "" ? itemHeight : "400px" }
                  : {}
                }
                    >
                      <h2
                        className="wcNote"
                        style={{
                          fontSize: headingSize !== "" ? headingSize : "42px",
                          color: headingColor !== "" ? headingColor : "#fff",
                          display:
                            headingShow === true ? "-webkit-box" : "none",
                        }}
                      >
                        {item.Title !== "" ? item.Title : "Heading"}
                      </h2>

                      <p
                        className="wcDesc"
                        style={{
                          fontSize:
                            descriptionSize !== "" ? descriptionSize : "18px",
                          color:
                            descriptionColor !== "" ? descriptionColor : "#fff",
                          display:
                            descriptionShow === true ? "-webkit-box" : "none",
                        }}
                        // dangerouslySetInnerHTML={{ __html: sanitize(item.Description ? item.Description : "Description"),}}
                      >
                        {item.Description !== "" ?
                          item.Description
                        : "Description"}
                      </p>

                      <a
                        href={item.ButtonLink?.Url || "#"}
                        target="_blank"
                        rel="noopener noreferrer"
                        data-interception="off"
                        style={{
                          fontSize: btnTextSize !== "" ? btnTextSize : "#fff",
                          backgroundColor:
                            btnBgColor !== "" ? btnBgColor : "#fff",
                          borderRadius:
                            btnBorderRadius !== "" ? btnBorderRadius : "0px",
                          color: BtnTextColor !== "" ? BtnTextColor : "#fff",
                          display: btnShow === true ? "block" : "none",
                        }}
                      >
                        {item.ButtonLink?.Description || "Button"}
                      </a>
                    </div>
                  )}
                </div>
              {/* </a> */}
            </SwiperSlide>
          ))}
        </Swiper>
      )}
    </>
  );
};

export default Lab02;
