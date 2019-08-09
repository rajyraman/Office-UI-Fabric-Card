import * as React from "react";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import { Persona, PersonaSize } from "office-ui-fabric-react/lib/Persona";
import { Card, ICardItemTokens, ICardTokens } from "@uifabric/react-cards";
import {
  FontWeights,
  mergeStyleSets,
  ColorClassNames
} from "@uifabric/styling";
import {
  Icon,
  Image,
  ImageFit,
  Stack,
  IStackTokens,
  IStackStyles,
  Text,
  ImageCoverStyle,
  ITextStyles
} from "office-ui-fabric-react";
import { DefaultPalette } from "office-ui-fabric-react/lib/Styling";

initializeIcons(undefined, { disableWarnings: true });

export interface IContactCardProps {
  body: string;
  bodyCaption: string;
  mainHeader: string;
  subHeader: string;
  cardImage: string;
  cardData: IContactCard[];
  triggerNavigate?: (id: string) => void;
}
export interface IAttributeValue {
  attribute: string;
  value: string;
}

export interface IContactCard {
  key: string;
  values: IAttributeValue[];
}

const getAttributeValue = (
  attributes: IAttributeValue[],
  attributeName: string
) => {
  if (attributes.findIndex(v => v.attribute == attributeName) == -1) return "";
  return attributes.find(v => v.attribute == attributeName)!.value;
};

export function ContactCard(props: IContactCardProps): JSX.Element {
  const styles = mergeStyleSets({
    descriptionText: {
      color: "#333333",
      padding: 10,
    },
    helpfulText: {
      color: "#333333",
      fontSize: 12,
      fontWeight: FontWeights.regular
    },
    icon: {
      color: "red",
      fontSize: 16,
      padding: 4,
      fontWeight: FontWeights.regular
    },
    cardItem: {
      minHeight: 260
    },
    persona: {
      padding: 5
    },
    caption: {
      textAlign: "center",
      fontWeight: FontWeights.semibold
    }
  });

  const sectionStackTokens: IStackTokens = { childrenGap: 20 };

  const stackStyles: IStackStyles = {
    root: {
      width: "100%",
      overflow: "auto"
    }
  };
  const captionStyles: ITextStyles = {
    root: {
      textAlign: 'center'
    }
  }
  const cardTokens: ICardTokens = {
    width: 400,
    padding: 20,
    height: 600,
    boxShadow: '0 0 20px rgba(0, 0, 0, .2)'
  }
  const cardClicked = (ev: React.MouseEvent<HTMLElement>): void => {
    if(props.triggerNavigate){
      props.triggerNavigate(ev.currentTarget.id);
    }
  };
  console.log(props.cardData);
  return (
    <Stack
      horizontal
      verticalFill
      tokens={sectionStackTokens}
      grow
      wrap
      styles={stackStyles}
    >
      {props.cardData.map(c => (
        <Card onClick={cardClicked} id={c.key} tokens={cardTokens}>
          <Persona
            text={getAttributeValue(c.values, props.mainHeader)}
            secondaryText={getAttributeValue(c.values, props.subHeader)}
            optionalText={getAttributeValue(c.values, props.subHeader)}
            size={PersonaSize.size56}
            className={styles.persona}
          />
          <Card.Item className={styles.cardItem}>
            <Image
              src={`data:image/jpg;base64,${getAttributeValue(
                c.values,
                props.cardImage
              )}`}
              imageFit={ImageFit.contain}
              maximizeFrame={true}
              coverStyle={ImageCoverStyle.portrait}
            />
          </Card.Item>
          <Text variant={"smallPlus"} className={styles.caption}>
            {getAttributeValue(c.values, props.bodyCaption)}
          </Text>          
          <Text className={styles.descriptionText} variant={"medium"}>
            {getAttributeValue(c.values, props.body)}
          </Text>
        </Card>
      ))}
    </Stack>
  );
}
