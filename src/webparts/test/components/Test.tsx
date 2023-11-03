import * as React from "react";
import { ITestProps } from "./ITestProps";
import {
  Stack,
  Persona,
  PersonaSize,
  DocumentCard,
  Text,
} from "@fluentui/react";
import styles from "./Test.module.scss";

export default class Test extends React.Component<ITestProps> {
  public render(): React.ReactElement<ITestProps> {
    console.log("Birthdays:", this.props.birthdays);

    // If no list is selected
    if (!this.props.selectedList) {
      return <div>Please select a list.</div>;
    }
    // If no data is being shown or an error occurred
    if (
      !Array.isArray(this.props.birthdays) ||
      this.props.birthdays.length === 0
    ) {
      return (
        <div>
          This list might not contain "Person" and "Birthdate" columns or an
          error occurred while fetching data.
        </div>
      );
    }

    return (
      <div
        className={`${styles.webPartContainer} ${styles[this.props.boxSize]}`}
      >
        <Stack
          tokens={{ childrenGap: 10 }}
          styles={{ root: { maxWidth: "100%" } }}
        >
          {this.props.birthdays.map((birthday, index) => {
            console.log("Birthday Person:", birthday.Person);
            const birthdayDate = new Date(birthday.Birthdate);
            const formattedDate = `${birthdayDate.getDate()}. ${birthdayDate.toLocaleString(
              "default",
              { month: "long" }
            )}`;

            // Check if the Department property exists and is not null or empty
            const department = birthday.Department ? (
              <Text variant="medium">{birthday.Department}</Text> // Change variant to "medium"
            ) : null;

            return (
              <DocumentCard
                key={index}
                className={`${styles.birthdayCard} ${
                  styles[this.props.boxSize as keyof typeof styles]
                }`}
              >
                <div className={styles.cardContentWrapper}>
                  <Stack horizontal className={styles.personaStack}>
                    <Persona
                      imageUrl={`${this.props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?UserName=${birthday.Person.EMail}&size=L`}
                      size={PersonaSize.size48}
                      className={styles.customPersona}
                    />
                    <Stack>
                      <Text variant="medium">{birthday.Person.Title}</Text>
                      {department} {/* Render department if it exists */}
                      <Text variant="medium">{formattedDate}</Text>
                    </Stack>
                  </Stack>
                </div>
              </DocumentCard>
            );
          })}
        </Stack>
      </div>
    );
  }
}
