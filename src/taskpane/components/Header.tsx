import * as React from "react";
import { makeStyles, Title3 } from "@fluentui/react-components";

interface HeaderProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    padding: "12px 16px",
    borderBottom: "1px solid #E1E1E1",
    backgroundColor: "#FFFFFF",
  },
  title: {
    margin: 0,
    fontWeight: 600,
  },
});

const Header: React.FC<HeaderProps> = (props: HeaderProps) => {
  const styles = useStyles();

  return (
    <header className={styles.root}>
      <Title3 className={styles.title}>{props.title}</Title3>
    </header>
  );
};

export default Header;
