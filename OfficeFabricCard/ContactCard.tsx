import * as React from "react";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import { Persona } from "office-ui-fabric-react/lib/Persona";
import { Card } from "@uifabric/react-cards";
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
  Text
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
      fontSize: 14,
      fontWeight: FontWeights.regular,
      padding: 10
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
    }
  });

  const tokens = {
    sectionStack: {
      childrenGap: 30
    },
    cardFooterStack: {
      childrenGap: 16
    }
  };
  const sectionStackTokens: IStackTokens = { childrenGap: 20 };

  const stackStyles: IStackStyles = {
    root: {
      width: "100%",
      overflow: "auto"
    }
  };

  const cardClicked = (ev: React.MouseEvent<HTMLElement>): void => {
    if(props.triggerNavigate){
      props.triggerNavigate(ev.currentTarget.id);
    }
  };

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
        <Card onClick={cardClicked} width={400} cellPadding={20} id={c.key}>
          <Persona
            text={getAttributeValue(c.values, props.mainHeader)}
            secondaryText={getAttributeValue(c.values, props.subHeader)}
          />
          <Card.Item>
            <Image
              src={`data:image/jpg;base64,${getAttributeValue(
                c.values,
                props.cardImage
              )}`}
              imageFit={ImageFit.contain}
              height={220}
            />
          </Card.Item>
          <Text className={styles.descriptionText}>
            {getAttributeValue(c.values, props.body)}
          </Text>
          <Card.Item>
            <Stack
              horizontal
              horizontalAlign="end"
              tokens={tokens.cardFooterStack}
              padding="12px 0 0"
              styles={{ root: { borderTop: "1px solid #F3F2F1" } }}
            >
              <Icon
                iconName="Heart"
                className={styles.icon}
                color={ColorClassNames.red}
              />
            </Stack>
          </Card.Item>
        </Card>
      ))}
    </Stack>
  );
}
