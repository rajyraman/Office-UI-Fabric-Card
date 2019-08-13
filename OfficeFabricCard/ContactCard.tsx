import * as React from "react";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";
import {
  CommandBar,
  ICommandBarItemProps
} from "office-ui-fabric-react/lib/CommandBar";
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
  ITextStyles,
  IImageStyles,
  TextStyles
} from "office-ui-fabric-react";
import {
  ScrollablePane,
  ScrollbarVisibility
} from "office-ui-fabric-react/lib/ScrollablePane";
import { Sticky, StickyPositionType } from "office-ui-fabric-react/lib/Sticky";

initializeIcons(undefined, { disableWarnings: true });

export interface IContactCardProps {
  body: string;
  bodyCaption: string;
  mainHeader: string;
  subHeader: string;
  cardImage: string;
  cardData: IContactCard[];
  layout?: string;
  totalResultCount: number;
  triggerNavigate?: (id: string) => void;
  triggerPaging?: (pageCommand: string) => void;
}
export interface IAttributeValue {
  attribute: string;
  value: string;
}

export interface IContactCard {
  key: string;
  values: IAttributeValue[];
}

export function ContactCard(props: IContactCardProps): JSX.Element {
  const styles = mergeStyleSets({
    descriptionText: {
      color: "#333333",
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
    },
    cardItem: {
    },
    persona: {
      padding: 5
    },
    caption: {
      textAlign: "center",
      fontWeight: FontWeights.semibold
    },
    imageStyle: {
      textAlign: "-webkit-center"
    },
    scrollableContainer:{
      position: "relative",
      minHeight: "70vh",
    }
  });

  const getAttributeValue = (
    attributes: IAttributeValue[],
    attributeName: string
  ) => {
    if (attributes.findIndex(v => v.attribute == attributeName) == -1)
      return "";
    return attributes.find(v => v.attribute == attributeName)!.value;
  };

  const stackTokens: IStackTokens = {
    childrenGap: 10,
    padding: 10
  };

  const stackStyles: IStackStyles = {
    root: {
      width: "100%",
      overflow: "auto",
      paddingTop: 10,
      paddingBottom: 10
    }
  };

  const cardTokens: ICardTokens = {
    padding: 20,
    boxShadow: "0 0 20px rgba(0, 0, 0, .2)"
  };

  const cardClicked = (ev: React.MouseEvent<HTMLElement>): void => {
    if (props.triggerNavigate) {
      props.triggerNavigate(ev.currentTarget.id);
    }
  };

  const rightCommands: ICommandBarItemProps[] = [
    {
      key: "next",
      name: `Load more (${props.cardData.length} of ${
        props.totalResultCount
      })..`,
      iconProps: {
        iconName: "ChevronRight"
      },
      disabled: props.cardData.length == props.totalResultCount,
      onClick: () => {
        if (props.triggerPaging) {
          props.triggerPaging("next");
        }
      }
    }
  ];
  const leftCommands: ICommandBarItemProps[] = [];

  return (
    <div className={styles.scrollableContainer}>
    <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
      <Sticky stickyPosition={StickyPositionType.Header}>
        <CommandBar farItems={rightCommands} items={leftCommands} />
      </Sticky>
      <Stack horizontal tokens={stackTokens} wrap styles={stackStyles}>
        {props.cardData.map(c => (
          <Card
            onClick={cardClicked}
            id={c.key}
            tokens={cardTokens}
            key={c.key}
            compact={props.layout == "compact"}
          >
            <Card.Item>
              <Persona
                text={getAttributeValue(c.values, props.mainHeader)}
                secondaryText={getAttributeValue(c.values, props.subHeader)}
                optionalText={getAttributeValue(c.values, props.subHeader)}
                size={
                  props.layout == "compact"
                    ? PersonaSize.small
                    : PersonaSize.size56
                }
                className={styles.persona}
              />
            </Card.Item>
            <Card.Item grow>
              <Image
                src={`data:image/jpg;base64,${getAttributeValue(
                  c.values,
                  props.cardImage
                )}`}
                shouldStartVisible={true}
                className={styles.imageStyle}
              />
            </Card.Item>
            <Stack>
              <Text variant={"smallPlus"} className={styles.caption}>
                {getAttributeValue(c.values, props.bodyCaption)}
              </Text>
              <Text className={styles.descriptionText} variant={"medium"}>
                {getAttributeValue(c.values, props.body)}
              </Text>
            </Stack>
          </Card>
        ))}
      </Stack>
    </ScrollablePane>
    </div>
  );
}
