import * as React from "react";
import { useState, useEffect } from "react";

// import styles from './Lab02.module.scss';
import type { ILab02Props } from "./ILab02Props";
import "./style.css";
import bannerBGImg from  "../assets/Bg.png";
// import { escape } from '@microsoft/sp-lodash-subset';

import { Swiper, SwiperSlide } from "swiper/react"; 

// import Swiper and modules styles
import "swiper/css";
import "swiper/css/navigation";
import "swiper/css/pagination";
import { Autoplay, Pagination } from "swiper"; 
 


const Lab02: React.FC<ILab02Props> = ({
  slider,
  listName,
  siteUrl,

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
}) => {
  const [heroSlider, setHeroSlider] = useState<boolean>(false);
  const [heroBanner, setHeroBanner] = useState<boolean>(true);

  const [items, setItems] = useState<any[]>([]);
  const apiUrl = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Title,Description,ButtonText,ButtonURL&$expand=AttachmentFiles`;


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
      document.documentElement.style.setProperty("--swiper-pagination-color", btnBgColor);
    }
  }, [btnBgColor]);

  useEffect(() =>{
    const fetchData = async () =>{
      try{
        const response = await fetch(apiUrl, {
          method: 'Get',
          headers: {
            'Accept': 'application/json;odata=verbose'
          }
        })
        const data:any = await response.json(); 

// Process items to add the full image URL if available
      const itemsWithImages = data.d.results.map((item: any) => {
        const imageUrl = item.AttachmentFiles.results.length > 0 ? `${item.AttachmentFiles.results[0].ServerRelativeUrl}` : null;
        return { ...item, imageUrl };
      });
      setItems(itemsWithImages);
        }
      catch(error){
        console.log('Error fetching data:', error);
      };  
    }
    fetchData();
  }, [siteUrl,listName]);
   
 console.log("data", items);
  return (
    <>
      {heroBanner && (
        <section
          className={`topBanner ${
            contentAlignment === "Right"
              ? "contRight"
              : contentAlignment === "Center"
              ? "contCenter"
              : "contLeft"
          } ${verticalAlign === "Top" ? "vertTop" : verticalAlign === "Bottom" ? "vertBottom" : "Center"}`}
          style={{minHeight:itemHeight !== "" ? itemHeight : "400px"}}
        >
          <div className="imgDiv" data-overlay={bgType}>
            <img
              src={bgImage}
              alt="BG image"
            />
          </div>
          <div className="wcContent">
            {/* <div className="borderCircle" >
                <img src="./profile.jpeg" alt="Contact Image">
            </div>  */}
            <h2
              className="wcNote"
              style={{
                fontSize: headingSize !== "" ? headingSize : "42px",
                color: headingColor !== "" ? headingColor : "#fff",
                display: headingShow === true ? "block" : "none",
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
              style={{
                fontSize: btnTextSize !== "" ? btnTextSize : "#fff",
                backgroundColor: btnBgColor !== "" ? btnBgColor : "#fff",
                borderRadius: btnBorderRadius !== "" ? btnBorderRadius : "0px",
                color: BtnTextColor !== "" ? BtnTextColor : "#fff",
                display: btnShow === true ? "block" : "none",
              }}
            >
              {btnText !== "" ? btnText : "Button Text"}
            </a>
          </div>
        </section>
      )}

      {heroSlider && (
        <Swiper
          spaceBetween={10} 
          loop = {true}
          speed={1200}
          autoplay={{
            delay:3000,
            disableOnInteraction:true,
          }}
          pagination={{
            clickable: true,
          }}
          modules={[Autoplay,Pagination]}
          className="mySwiper"
        >
          { items.map((item,index) => (
          <SwiperSlide key={index}>
            <div
              className={`topBanner ${
                contentAlignment === "Right" 
                  ? "contRight"
                  : contentAlignment === "Center"
                  ? "contCenter"
                  : "contLeft"
              } ${verticalAlign === "Top" ? "vertTop" : verticalAlign === "Bottom" ? "vertBottom" : "Center"}`}
              style={{minHeight:itemHeight !== "" ? itemHeight : "400px"}}
            >
              <div className="imgDiv" data-overlay={bgType}>
                <img
                  src={item.imageUrl || bgImage}
                  alt="BG image"
                />
              </div>
              <div className="wcContent"> 
                <h2
                  className="wcNote"
                  style={{
                    fontSize: headingSize !== "" ? headingSize : "42px",
                    color: headingColor !== "" ? headingColor : "#fff",
                    display: headingShow === true ? "block" : "none",
                  }}
                >
                  {item.Title !== "" ? item.Title : "Heading"}
                </h2>

                <p
                  style={{
                    fontSize: descriptionSize !== "" ? descriptionSize : "18px",
                    color: descriptionColor !== "" ? descriptionColor : "#fff",
                    display: descriptionShow === true ? "-webkit-box" : "none",
                  }}
                >
                  {item.Description !== "" ? item.Description : "Description"}
                </p>

                <a
                  href={item.ButtonURL || '#' }
                  style={{
                    fontSize: btnTextSize !== "" ? btnTextSize : "#fff",
                    backgroundColor: btnBgColor !== "" ? btnBgColor : "#fff",
                    borderRadius: btnBorderRadius !== "" ? btnBorderRadius : "0px",
                    color: BtnTextColor !== "" ? BtnTextColor : "#fff",
                    display: btnShow === true ? "block" : "none",
                  }}
                >
                  {item.ButtonText !== "" ? item.ButtonText : "Button"}
                </a>
              </div>
            </div>
          </SwiperSlide>
))}
        </Swiper>
      )}
    </>
  );
};

export default Lab02;
